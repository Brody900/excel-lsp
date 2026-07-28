"""Workbook-and-index orchestration for surgical write tools."""

from __future__ import annotations

import hashlib
import os
from collections.abc import Mapping, Sequence
from datetime import UTC, datetime
from pathlib import Path

from excel_lsp.core.edit.models import (
    CellEdit,
    CellEditKind,
    ColumnFormulaResult,
    EditResult,
    PatchResult,
)
from excel_lsp.core.edit.writer import patch_workbook
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.formulas.a1 import CellRef
from excel_lsp.core.formulas.from_r1c1 import from_r1c1
from excel_lsp.core.formulas.translation import translate_a1_formula
from excel_lsp.core.index.lifecycle import index_workbook
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import CellRecord, Rect, SheetDescriptor, SheetParseSummary
from excel_lsp.core.parse import OOXMLParser
from excel_lsp.core.parse.coordinates import parse_cell_ref
from excel_lsp.core.regions import RegionOptions


def write_cells(
    path: str | Path,
    edits: Sequence[CellEdit],
    *,
    index_dir: str | Path | None = None,
    _expected_generation: int | None = None,
    _max_cells: int = 500,
    _written_rectangles: Sequence[tuple[str, Rect]] | None = None,
    _formula_rectangles: Sequence[tuple[str, Rect]] | None = None,
) -> EditResult:
    """Surgically edit cells, patch the index, and mark transitive staleness."""
    workbook = Path(path).expanduser().resolve()
    index_update = index_workbook(workbook, index_dir=index_dir)
    edit_rectangles, edit_formula_rectangles = _edit_rectangles(edits)
    written_rectangles = tuple(_written_rectangles or edit_rectangles)
    formula_rectangles = tuple(_formula_rectangles or edit_formula_rectangles)
    patch_result: PatchResult | None = None
    stale_rectangles: tuple[tuple[str, Rect], ...] = ()
    timestamp = datetime.now(UTC).isoformat()

    try:
        with IndexStore(index_update.index_path) as store, store.transaction():
            if _expected_generation is not None and store.generation != _expected_generation:
                raise ExcelLSPError(
                    ErrorCode.CONFLICT,
                    "Column symbol changed while its edit was being prepared.",
                    hint="Resolve the column symbol again and retry.",
                )
            expected_hash = store.get_meta("workbook_hash")
            if expected_hash is None:
                raise ExcelLSPError(
                    ErrorCode.CONFLICT,
                    "Workbook index has no source hash.",
                    hint="Run refresh and retry the edit.",
                )
            stale_rectangles = store.plan_staleness(
                written_rectangles,
                formula_rectangles=formula_rectangles,
            )
            region_options = _stored_region_options(store)
            patch_result = patch_workbook(
                workbook,
                edits,
                expected_workbook_hash=expected_hash,
                _max_cells=_max_cells,
            )
            with OOXMLParser(workbook) as parser:
                if parser.hashes.whole_file != patch_result.workbook_hash_after:
                    raise ExcelLSPError(
                        ErrorCode.CONFLICT,
                        "Workbook changed immediately after the edit was installed.",
                        hint="Review the workbook, run refresh, and retry if needed.",
                    )
                collection_stat = _verify_workbook_snapshot(
                    workbook,
                    patch_result.workbook_hash_after,
                )
                try:
                    sheet_patches = _collect_sheet_patches(parser, patch_result)
                except Exception as collection_error:
                    try:
                        _verify_workbook_snapshot(
                            workbook,
                            patch_result.workbook_hash_after,
                            expected_stat=collection_stat,
                        )
                    except ExcelLSPError as conflict:
                        raise conflict from collection_error
                    raise
                file_stat = _verify_workbook_snapshot(
                    workbook,
                    patch_result.workbook_hash_after,
                    expected_stat=collection_stat,
                )
                generation = store.apply_editor_patch(
                    parser.metadata,
                    sheet_patches,
                    styles=parser.styles,
                    package_hashes=parser.hashes,
                    expected_workbook_hash=expected_hash,
                    mtime_ns=file_stat.st_mtime_ns,
                    size=file_stat.st_size,
                    indexed_at=timestamp,
                    stale_rectangles=stale_rectangles,
                    stale_since=timestamp,
                    region_options=region_options,
                )
                _verify_workbook_snapshot(
                    workbook,
                    patch_result.workbook_hash_after,
                    expected_stat=file_stat,
                )
        _verify_workbook_snapshot(
            workbook,
            patch_result.workbook_hash_after,
            expected_stat=file_stat,
        )
        return EditResult(
            patch=patch_result,
            generation=generation,
            stale_blocks=len(stale_rectangles),
            direct_index_patch=True,
        )
    except Exception as primary_error:
        if patch_result is None:
            raise
        recovered = _recover_written_workbook(
            workbook,
            index_dir,
            patch_result,
            stale_rectangles,
            timestamp,
            primary_error,
        )
        if isinstance(primary_error, ExcelLSPError) and primary_error.code is ErrorCode.CONFLICT:
            primary_error.add_note(
                "The workbook reached replacement, and its sidecar was reconciled before "
                "this conflict was returned."
            )
            raise primary_error
        return recovered


def set_column_formula(
    path: str | Path,
    column_symbol: str,
    formula: str,
    *,
    overwrite: bool = False,
    index_dir: str | Path | None = None,
) -> ColumnFormulaResult:
    """Fill one semantic column body from an A1-anchor or R1C1 formula."""
    if not formula.startswith("=") or len(formula) == 1:
        raise ExcelLSPError(
            ErrorCode.INVALID_VALUE,
            "Column formula must be a nonempty string beginning with '='.",
        )
    workbook = Path(path).expanduser().resolve()
    opened = index_workbook(workbook, index_dir=index_dir)
    with IndexStore(opened.index_path) as store:
        sheet, target, occupied = store.resolve_column_write_target(column_symbol)
        generation = store.generation
    if occupied and not overwrite:
        raise ExcelLSPError(
            ErrorCode.INVALID_VALUE,
            f"Column {column_symbol!r} already contains {occupied} indexed cells.",
            hint="Pass overwrite=true only after reviewing the existing column.",
        )

    anchor = CellRef(target.row_min, target.col_min)
    try:
        anchor_r1c1 = from_r1c1(formula, anchor)
        if anchor_r1c1 is None:
            formulas = tuple(
                translate_a1_formula(
                    formula,
                    origin=anchor,
                    target=CellRef(row, target.col_min),
                )
                for row in range(target.row_min, target.row_max + 1)
            )
        else:
            formulas = tuple(
                from_r1c1(formula, CellRef(row, target.col_min))
                for row in range(target.row_min, target.row_max + 1)
            )
            if any(rendered is None for rendered in formulas):
                raise ValueError("R1C1 formula classification changed between column rows")
    except ValueError as exc:
        raise ExcelLSPError(
            ErrorCode.INVALID_VALUE,
            "Column formula cannot be translated across the complete body range.",
            hint="Check its reference mode and boundary-relative references.",
        ) from exc

    edits = tuple(
        CellEdit.formula(sheet, f"{_column_label(target.col_min)}{row}", rendered)
        for row, rendered in zip(
            range(target.row_min, target.row_max + 1),
            formulas,
            strict=True,
        )
        if rendered is not None
    )
    edit_result = write_cells(
        workbook,
        edits,
        index_dir=index_dir,
        _expected_generation=generation,
        _max_cells=len(edits),
        _written_rectangles=((sheet, target),),
        _formula_rectangles=((sheet, target),),
    )
    with IndexStore(opened.index_path) as store:
        block = store.formula_block_at(sheet, target.row_min, target.col_min)
    return ColumnFormulaResult(
        edit=edit_result,
        formula_block=block,
        cells_written=len(edits),
    )


def _column_label(column: int) -> str:
    from excel_lsp.core.parse.coordinates import column_label

    return column_label(column)


def _edit_rectangles(
    edits: Sequence[CellEdit],
) -> tuple[tuple[tuple[str, Rect], ...], tuple[tuple[str, Rect], ...]]:
    written: list[tuple[str, Rect]] = []
    formulas: list[tuple[str, Rect]] = []
    for edit in edits:
        try:
            row, col = parse_cell_ref(edit.ref)
        except ValueError as exc:
            raise ExcelLSPError(
                ErrorCode.INVALID_REF,
                f"Invalid cell reference: {edit.ref!r}.",
            ) from exc
        region = (edit.sheet, Rect(row, row, col, col))
        written.append(region)
        if edit.kind is CellEditKind.FORMULA:
            formulas.append(region)
    return tuple(written), tuple(formulas)


def _stored_region_options(store: IndexStore) -> RegionOptions:
    raw = store.get_meta("region_gap_tol")
    try:
        return RegionOptions(gap_tol=int(raw or ""))
    except ValueError as exc:
        raise ExcelLSPError(
            ErrorCode.CORRUPT,
            "Index contains an invalid region-analysis configuration.",
        ) from exc


def _collect_sheet_patches(
    parser: OOXMLParser,
    patch_result: PatchResult,
) -> Mapping[
    SheetDescriptor,
    tuple[SheetParseSummary, Mapping[tuple[int, int], CellRecord | None]],
]:
    targets_by_sheet: dict[str, set[tuple[int, int]]] = {}
    for patched_cell in patch_result.patched_cells:
        targets_by_sheet.setdefault(patched_cell.sheet, set()).add(parse_cell_ref(patched_cell.ref))

    descriptors = {descriptor.name: descriptor for descriptor in parser.metadata.sheets}
    result: dict[
        SheetDescriptor,
        tuple[SheetParseSummary, Mapping[tuple[int, int], CellRecord | None]],
    ] = {}
    for sheet_name in sorted(targets_by_sheet):
        descriptor = descriptors.get(sheet_name)
        if descriptor is None:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                f"Patched worksheet is absent from the package: {sheet_name!r}.",
            )
        targets = targets_by_sheet[sheet_name]
        cells: dict[tuple[int, int], CellRecord | None] = {target: None for target in targets}

        def collect(
            cell: CellRecord,
            current_targets: set[tuple[int, int]] = targets,
            current_cells: dict[tuple[int, int], CellRecord | None] = cells,
        ) -> None:
            key = (cell.row, cell.col)
            if key in current_targets:
                current_cells[key] = cell

        summary = parser.parse_sheet(descriptor, collect)
        result[descriptor] = (summary, cells)
    return result


def _verify_workbook_snapshot(
    workbook: Path,
    expected_hash: str,
    *,
    expected_stat: os.stat_result | None = None,
) -> os.stat_result:
    """Prove that a path hash and stat describe one unchanged file generation."""
    try:
        stat_before = workbook.stat()
        digest = hashlib.sha256()
        with workbook.open("rb") as stream:
            while chunk := stream.read(1024 * 1024):
                digest.update(chunk)
        stat_after = workbook.stat()
    except OSError as exc:
        raise ExcelLSPError(
            ErrorCode.CONFLICT,
            "Workbook changed while its edited index was being prepared.",
            hint="Review the workbook, run refresh, and retry if needed.",
        ) from exc

    before_fingerprint = _stat_fingerprint(stat_before)
    after_fingerprint = _stat_fingerprint(stat_after)
    if (
        before_fingerprint != after_fingerprint
        or digest.hexdigest() != expected_hash
        or (expected_stat is not None and after_fingerprint != _stat_fingerprint(expected_stat))
    ):
        raise ExcelLSPError(
            ErrorCode.CONFLICT,
            "Workbook changed while its edited index was being prepared.",
            hint="Review the workbook, run refresh, and retry if needed.",
        )
    return stat_after


def _stat_fingerprint(value: os.stat_result) -> tuple[int, int, int, int, int]:
    return (
        value.st_dev,
        value.st_ino,
        value.st_size,
        value.st_mtime_ns,
        value.st_ctime_ns,
    )


def _recover_written_workbook(
    workbook: Path,
    index_dir: str | Path | None,
    patch_result: PatchResult,
    stale_rectangles: Sequence[tuple[str, Rect]],
    timestamp: str,
    primary_error: Exception,
) -> EditResult:
    """Rebuild a sidecar if a post-replacement direct patch unexpectedly fails."""
    try:
        update = index_workbook(workbook, index_dir=index_dir)
        with IndexStore(update.index_path) as store:
            generation = store.record_staleness(stale_rectangles, since=timestamp)
    except Exception as recovery_error:
        failure = ExcelLSPError(
            ErrorCode.INTERNAL,
            "Workbook edit succeeded, but its index could not be recovered.",
            hint="Run refresh before issuing another workbook operation.",
            details={"path": str(workbook)},
        )
        failure.add_note(f"Direct index patch failed: {type(primary_error).__name__}.")
        failure.add_note(f"Index recovery failed: {type(recovery_error).__name__}.")
        raise failure from recovery_error
    return EditResult(
        patch=patch_result,
        generation=generation,
        stale_blocks=len(stale_rectangles),
        direct_index_patch=False,
    )


__all__ = ["set_column_formula", "write_cells"]
