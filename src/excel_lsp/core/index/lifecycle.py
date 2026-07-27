"""Hash-aware workbook index placement and freshness lifecycle."""

from __future__ import annotations

import hashlib
import json
import os
from datetime import UTC, datetime
from pathlib import Path

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import IndexUpdate, PackageHashes
from excel_lsp.core.parse import OOXMLParser
from excel_lsp.core.regions import RegionOptions

_SUPPORTED_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})
_WORKBOOK_PART = "xl/workbook.xml"
_WORKBOOK_RELS_PART = "xl/_rels/workbook.xml.rels"
_CONTENT_TYPES_PART = "[Content_Types].xml"
_SHARED_STRINGS_PART = "xl/sharedStrings.xml"
_STYLES_PART = "xl/styles.xml"
_WORKBOOK_STRUCTURE_PARTS = frozenset(
    {
        _WORKBOOK_PART,
        _WORKBOOK_RELS_PART,
        _CONTENT_TYPES_PART,
    }
)
_GLOBAL_VALUE_PARTS = frozenset({_SHARED_STRINGS_PART, _STYLES_PART})
_FORMULA_CONTEXT_PREFIXES = (
    "xl/externalLinks/",
    "xl/tables/",
    "xl/worksheets/_rels/",
)
_REGION_ANALYSIS_VERSION = "1"
_FORMULA_ANALYSIS_VERSION = "1"


class _WorkbookChangedDuringIndex(RuntimeError):
    pass


def resolve_index_path(
    workbook_path: str | Path,
    index_dir: str | Path | None = None,
) -> Path:
    """Resolve the next-to-file sidecar or path-hashed shared-index location."""
    workbook = Path(workbook_path).expanduser().resolve()
    configured_dir: str | Path | None = index_dir
    if configured_dir is None:
        configured_dir = os.environ.get("EXCEL_LSP_INDEX_DIR") or None
    if configured_dir is None:
        return workbook.with_name(f"{workbook.name}.xlsp.db")

    root = Path(configured_dir).expanduser().resolve()
    digest = hashlib.sha256(str(workbook).encode("utf-8")).hexdigest()[:8]
    return root / f"{workbook.stem}.{digest}.xlsp.db"


def index_workbook(
    path: str | Path,
    *,
    index_dir: str | Path | None = None,
    gap_tol: int | None = None,
) -> IndexUpdate:
    """Open or refresh a workbook index using stat and selected-part hashes."""
    requested_region_options = None if gap_tol is None else RegionOptions(gap_tol=gap_tol)
    workbook = Path(path).expanduser().resolve()
    initial_stat = _stat_workbook(workbook)
    index_path = resolve_index_path(workbook, index_dir)

    with IndexStore(index_path) as store:
        region_options = requested_region_options or _preserved_region_options(store, workbook)
        if not store.schema_rebuilt and _fast_path_matches(
            store,
            workbook,
            initial_stat,
            region_options,
        ):
            return IndexUpdate(
                workbook_path=str(workbook),
                index_path=str(index_path),
                generation=store.generation,
                changed=False,
                reindexed_sheets=(),
            )

        source_stat = initial_stat
        for attempt in range(2):
            try:
                return _index_from_parser(workbook, source_stat, store, region_options)
            except _WorkbookChangedDuringIndex as exc:
                if attempt:
                    raise ExcelLSPError(
                        ErrorCode.CORRUPT,
                        "workbook changed repeatedly while it was being indexed",
                        hint="Wait for Excel to finish saving, then retry refresh.",
                    ) from exc
                source_stat = _stat_workbook(workbook)
            except ExcelLSPError as exc:
                if exc.code is not ErrorCode.CORRUPT:
                    raise
                current_stat = _stat_workbook(workbook)
                if attempt or _same_stat(source_stat, current_stat):
                    raise
                source_stat = current_stat
        raise AssertionError("unreachable lifecycle retry state")


def ensure_fresh(
    path: str | Path,
    *,
    index_dir: str | Path | None = None,
    gap_tol: int | None = None,
) -> IndexUpdate:
    """Ensure the sidecar matches the workbook before serving a core API call."""
    return index_workbook(path, index_dir=index_dir, gap_tol=gap_tol)


def _index_from_parser(
    workbook: Path,
    source_stat: os.stat_result,
    store: IndexStore,
    region_options: RegionOptions,
) -> IndexUpdate:
    with OOXMLParser(workbook) as parser:
        metadata = parser.metadata
        hashes = parser.hashes
        old_parts = store.get_part_hashes()
        old_workbook_hash = store.get_meta("workbook_hash")
        stored_path = store.get_meta("workbook_path")
        full_rebuild = (
            store.schema_rebuilt
            or old_workbook_hash is None
            or stored_path is None
            or not _paths_equal(stored_path, workbook)
            or store.get_meta("analysis_version") != _REGION_ANALYSIS_VERSION
            or store.get_meta("formula_analysis_version") != _FORMULA_ANALYSIS_VERSION
            or store.get_meta("region_gap_tol") != str(region_options.gap_tol)
        )
        if not full_rebuild and old_workbook_hash == hashes.whole_file:
            ending_stat = _stat_workbook(workbook)
            if not _same_stat(source_stat, ending_stat):
                raise _WorkbookChangedDuringIndex
            with store.transaction():
                store.set_meta_many(
                    {
                        "mtime_ns": ending_stat.st_mtime_ns,
                        "size": ending_stat.st_size,
                    }
                )
            return IndexUpdate(
                workbook_path=str(workbook),
                index_path=str(store.path),
                generation=store.generation,
                changed=False,
                reindexed_sheets=(),
            )
        changed_parts = _changed_parts(old_parts, hashes)
        workbook_structure_changed = bool(changed_parts.intersection(_WORKBOOK_STRUCTURE_PARTS))
        global_values_changed = bool(changed_parts.intersection(_GLOBAL_VALUE_PARTS))
        formula_context_changed = (
            full_rebuild
            or workbook_structure_changed
            or any(part.startswith(_FORMULA_CONTEXT_PREFIXES) for part in changed_parts)
        )

        if full_rebuild or workbook_structure_changed or global_values_changed:
            sheets_to_reindex = metadata.sheets
        else:
            sheets_to_reindex = tuple(
                descriptor
                for descriptor in metadata.sheets
                if any(
                    _normalize_part_name(part) in changed_parts
                    for part in (descriptor.xml_part, *descriptor.related_parts)
                )
            )

        with store.transaction():
            if full_rebuild or workbook_structure_changed:
                store.replace_sheet_catalog(metadata.sheets)
                store.replace_defined_names(metadata)

            store.prepare_list_object_refresh(sheets_to_reindex)
            for descriptor in sheets_to_reindex:
                store.replace_sheet(
                    descriptor,
                    lambda on_cell, sheet=descriptor: parser.parse_sheet(sheet, on_cell),
                    styles=parser.styles,
                    region_options=region_options,
                )

            formula_sheets = metadata.sheets if formula_context_changed else sheets_to_reindex
            store.replace_formula_analysis(metadata, formula_sheets)

            ending_stat = _stat_workbook(workbook)
            if not _same_stat(source_stat, ending_stat):
                raise _WorkbookChangedDuringIndex

            part_kinds = {
                _normalize_part_name(descriptor.xml_part): descriptor.kind
                for descriptor in metadata.sheets
            }
            for descriptor in metadata.sheets:
                for part in descriptor.related_parts:
                    part_kinds[_normalize_part_name(part)] = "worksheet_metadata"
            store.replace_part_hashes(hashes.parts, kinds=part_kinds)
            store.set_meta_many(
                {
                    "schema_version": store.get_meta("schema_version", "1") or "1",
                    "workbook_path": str(workbook),
                    "workbook_hash": hashes.whole_file,
                    "indexed_at": datetime.now(UTC).isoformat(),
                    "mtime_ns": ending_stat.st_mtime_ns,
                    "size": ending_stat.st_size,
                    "date1904": int(metadata.date1904),
                    "has_vba": int(metadata.has_vba),
                    "external_links": json.dumps(
                        dict(sorted(metadata.external_links.items())),
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ),
                    "analysis_version": _REGION_ANALYSIS_VERSION,
                    "formula_analysis_version": _FORMULA_ANALYSIS_VERSION,
                    "region_gap_tol": region_options.gap_tol,
                }
            )
            generation = store.bump_generation()

        return IndexUpdate(
            workbook_path=str(workbook),
            index_path=str(store.path),
            generation=generation,
            changed=True,
            reindexed_sheets=tuple(sheet.name for sheet in sheets_to_reindex),
        )


def _fast_path_matches(
    store: IndexStore,
    workbook: Path,
    file_stat: os.stat_result,
    region_options: RegionOptions,
) -> bool:
    stored_path = store.get_meta("workbook_path")
    if stored_path is None or not _paths_equal(stored_path, workbook):
        return False
    try:
        return (
            int(store.get_meta("mtime_ns", "-1") or "-1") == file_stat.st_mtime_ns
            and int(store.get_meta("size", "-1") or "-1") == file_stat.st_size
            and store.get_meta("workbook_hash") is not None
            and store.get_meta("analysis_version") == _REGION_ANALYSIS_VERSION
            and store.get_meta("formula_analysis_version") == _FORMULA_ANALYSIS_VERSION
            and store.get_meta("region_gap_tol") == str(region_options.gap_tol)
        )
    except ValueError:
        return False


def _changed_parts(old_parts: dict[str, str], hashes: PackageHashes) -> set[str]:
    new_parts = {_normalize_part_name(name): part_hash for name, part_hash in hashes.parts.items()}
    names = old_parts.keys() | new_parts.keys()
    return {name for name in names if old_parts.get(name) != new_parts.get(name)}


def _stat_workbook(workbook: Path) -> os.stat_result:
    try:
        file_stat = workbook.stat()
    except FileNotFoundError as exc:
        raise ExcelLSPError(
            ErrorCode.NOT_FOUND,
            f"workbook not found: {workbook}",
        ) from exc
    except PermissionError as exc:
        raise ExcelLSPError(
            ErrorCode.LOCKED,
            f"workbook cannot be read: {workbook}",
            hint="Close the application holding the file and retry.",
        ) from exc
    if not workbook.is_file():
        raise ExcelLSPError(ErrorCode.NOT_FOUND, f"workbook is not a file: {workbook}")
    if workbook.suffix.casefold() not in _SUPPORTED_SUFFIXES:
        raise ExcelLSPError(
            ErrorCode.UNSUPPORTED_FORMAT,
            f"unsupported workbook format: {workbook.suffix or '<none>'}",
            hint="Use .xlsx, .xlsm, .xltx, or .xltm.",
        )
    return file_stat


def _paths_equal(stored_path: str, workbook: Path) -> bool:
    try:
        resolved_stored_path = Path(stored_path).resolve()
    except (OSError, RuntimeError, ValueError):
        return False
    return os.path.normcase(str(resolved_stored_path)) == os.path.normcase(str(workbook))


def _preserved_region_options(store: IndexStore, workbook: Path) -> RegionOptions:
    """Reuse a valid same-workbook tolerance, otherwise use the documented default."""
    stored_path = store.get_meta("workbook_path")
    if stored_path is None or not _paths_equal(stored_path, workbook):
        return RegionOptions()
    raw_gap_tol = store.get_meta("region_gap_tol")
    try:
        return RegionOptions(gap_tol=int(raw_gap_tol or ""))
    except ValueError:
        return RegionOptions()


def _same_stat(left: os.stat_result, right: os.stat_result) -> bool:
    return left.st_mtime_ns == right.st_mtime_ns and left.st_size == right.st_size


def _normalize_part_name(part_name: str) -> str:
    return part_name.replace("\\", "/").lstrip("/")


__all__ = ["ensure_fresh", "index_workbook", "resolve_index_path"]
