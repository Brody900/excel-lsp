"""Surgical OOXML workbook patching with byte-identical untouched parts."""

# pyright: reportAttributeAccessIssue=false, reportMissingTypeStubs=false
# pyright: reportUnknownArgumentType=false, reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false
# pyright: reportUnknownVariableType=false

from __future__ import annotations

import hashlib
import math
import os
import posixpath
import tempfile
import time
from collections.abc import Iterable, Mapping, Sequence
from contextlib import suppress
from datetime import date, datetime
from datetime import time as datetime_time
from pathlib import Path
from zipfile import BadZipFile, LargeZipFile, ZipFile, ZipInfo

from lxml import etree

from excel_lsp.core.edit.models import CellEdit, CellEditKind, PatchedCell, PatchResult
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.formulas.a1 import CellRef
from excel_lsp.core.formulas.translation import translate_a1_formula
from excel_lsp.core.models import Rect
from excel_lsp.core.parse import OOXMLParser
from excel_lsp.core.parse._xml import attr_by_local, child_by_local, local_name, parse_xml
from excel_lsp.core.parse.coordinates import make_cell_ref, parse_cell_ref, parse_rect

_CALC_CHAIN_PART = "xl/calcChain.xml"
_CONTENT_TYPES_PART = "[Content_Types].xml"
_WORKBOOK_PART = "xl/workbook.xml"
_WORKBOOK_RELS_PART = "xl/_rels/workbook.xml.rels"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_REPLACE_RETRY_DELAY_SECONDS = 0.5
_MAX_WRITE_CELLS = 500
_MAX_CELL_TEXT_LENGTH = 32_767
_MAX_FORMULA_LENGTH = 8_192
_CALC_PR_SUCCESSORS = frozenset(
    {
        "oleSize",
        "customWorkbookViews",
        "pivotCaches",
        "smartTagPr",
        "smartTagTypes",
        "webPublishing",
        "fileRecoveryPr",
        "webPublishObjects",
        "extLst",
    }
)


def patch_workbook(
    path: str | Path,
    edits: Sequence[CellEdit],
    *,
    expected_workbook_hash: str,
    _max_cells: int = _MAX_WRITE_CELLS,
) -> PatchResult:
    """Apply cell edits by replacing only deliberate OOXML package parts.

    The caller supplies the whole-file hash captured at index time.  The
    sibling Excel lockfile check deliberately precedes conflict detection.
    """
    workbook = Path(path).expanduser().resolve()
    _check_excel_lockfile(workbook)
    normalized_edits = _validate_edits(edits, max_cells=_max_cells)

    with OOXMLParser(workbook) as parser:
        metadata = parser.metadata
        actual_hash = parser.hashes.whole_file
    if actual_hash != expected_workbook_hash:
        raise ExcelLSPError(
            ErrorCode.CONFLICT,
            "Workbook changed since it was indexed.",
            hint="Run refresh, review the changed workbook, and retry the edit.",
        )

    descriptors = {descriptor.name: descriptor for descriptor in metadata.sheets}
    edits_by_part: dict[str, list[CellEdit]] = {}
    part_sheet_names: dict[str, str] = {}
    for edit in normalized_edits:
        descriptor = descriptors.get(edit.sheet)
        if descriptor is None:
            raise ExcelLSPError(
                ErrorCode.INVALID_REF,
                f"Unknown worksheet in edit reference: {edit.sheet!r}.",
            )
        if descriptor.kind != "worksheet":
            raise ExcelLSPError(
                ErrorCode.INVALID_REF,
                f"Sheet {edit.sheet!r} does not contain editable worksheet cells.",
            )
        part = _normalize_part_name(descriptor.xml_part)
        edits_by_part.setdefault(part, []).append(edit)
        part_sheet_names[part] = edit.sheet

    temporary_path: Path | None = None
    try:
        with ZipFile(workbook, "r") as source:
            infos = _index_archive(source)
            modified_payloads: dict[str, bytes] = {}
            patched_cells: list[PatchedCell] = []

            for part in sorted(edits_by_part):
                info = infos.get(part)
                if info is None:
                    raise _corrupt(f"worksheet package part is missing: {part}")
                original = _read_member(source, info)
                payload, cells = _patch_sheet_xml(
                    original,
                    part_sheet_names[part],
                    edits_by_part[part],
                )
                if payload != original:
                    modified_payloads[part] = payload
                patched_cells.extend(cells)

            workbook_info = infos.get(_WORKBOOK_PART)
            if workbook_info is None:
                raise _corrupt("workbook package part is missing")
            workbook_original = _read_member(source, workbook_info)
            workbook_payload = _ensure_full_calculation(workbook_original)
            if workbook_payload != workbook_original:
                modified_payloads[_WORKBOOK_PART] = workbook_payload

            deleted_parts = {_CALC_CHAIN_PART} if _CALC_CHAIN_PART in infos else set()
            _patch_calc_chain_references(source, infos, modified_payloads, deleted_parts)

            temporary_path = _write_replacement_archive(
                workbook,
                source,
                modified_payloads,
                deleted_parts,
            )

        if _hash_file(workbook) != expected_workbook_hash:
            raise ExcelLSPError(
                ErrorCode.CONFLICT,
                "Workbook changed while the edit was being prepared.",
                hint="Run refresh, review the changed workbook, and retry the edit.",
            )
        final_hash = _validate_replacement(temporary_path, part_sheet_names.values())
        _atomic_replace_with_retry(
            temporary_path,
            workbook,
            expected_workbook_hash=expected_workbook_hash,
        )
        temporary_path = None
        return PatchResult(
            path=workbook,
            workbook_hash_before=expected_workbook_hash,
            workbook_hash_after=final_hash,
            modified_parts=tuple(sorted(modified_payloads)),
            deleted_parts=tuple(sorted(deleted_parts)),
            patched_cells=tuple(
                sorted(
                    patched_cells,
                    key=lambda cell: (
                        cell.sheet.casefold(),
                        parse_cell_ref(cell.ref),
                        not cell.requested,
                    ),
                )
            ),
        )
    except ExcelLSPError:
        raise
    except (BadZipFile, LargeZipFile, etree.XMLSyntaxError, OSError, RuntimeError) as exc:
        raise _corrupt("Workbook could not be patched as a valid OOXML package.") from exc
    finally:
        if temporary_path is not None:
            with suppress(OSError):
                temporary_path.unlink(missing_ok=True)


def _validate_edits(
    edits: Sequence[CellEdit],
    *,
    max_cells: int,
) -> tuple[CellEdit, ...]:
    if type(max_cells) is not int or not 1 <= max_cells <= 1_048_576:
        raise ValueError("max_cells must be between 1 and 1,048,576")
    if not edits:
        raise ExcelLSPError(ErrorCode.INVALID_VALUE, "At least one cell edit is required.")
    if len(edits) > max_cells:
        raise ExcelLSPError(
            ErrorCode.INVALID_VALUE,
            f"A single write may contain at most {max_cells} cells.",
        )
    result: list[CellEdit] = []
    seen: set[tuple[str, int, int]] = set()
    for edit in edits:
        if not edit.sheet or not edit.sheet.strip():
            raise ExcelLSPError(ErrorCode.INVALID_REF, "Cell edits require a worksheet name.")
        try:
            row, col = parse_cell_ref(edit.ref)
        except ValueError as exc:
            raise ExcelLSPError(
                ErrorCode.INVALID_REF,
                f"Invalid cell reference: {edit.ref!r}.",
            ) from exc
        key = (edit.sheet.casefold(), row, col)
        if key in seen:
            raise ExcelLSPError(
                ErrorCode.INVALID_REF,
                f"Duplicate cell edit: {edit.sheet}!{make_cell_ref(row, col)}.",
            )
        seen.add(key)
        canonical_ref = make_cell_ref(row, col)
        if edit.kind is CellEditKind.FORMULA:
            if not isinstance(edit.payload, str) or not edit.payload.startswith("="):
                raise ExcelLSPError(
                    ErrorCode.INVALID_VALUE,
                    f"Formula write at {edit.sheet}!{canonical_ref} must start with '='.",
                )
            if len(edit.payload) == 1:
                raise ExcelLSPError(
                    ErrorCode.INVALID_VALUE,
                    f"Formula write at {edit.sheet}!{canonical_ref} is empty.",
                )
            formula_length = _utf16_code_units(
                edit.payload,
                sheet=edit.sheet,
                ref=canonical_ref,
                kind="Formula",
            )
            if formula_length > _MAX_FORMULA_LENGTH or edit.payload.startswith("=="):
                raise ExcelLSPError(
                    ErrorCode.INVALID_VALUE,
                    f"Formula write at {edit.sheet}!{canonical_ref} is not valid Excel text.",
                    hint=f"Use one leading '=' and at most {_MAX_FORMULA_LENGTH} characters.",
                )
        elif edit.kind is CellEditKind.VALUE:
            _validate_write_value(edit.payload, edit.sheet, canonical_ref)
        else:
            raise ExcelLSPError(ErrorCode.INVALID_VALUE, "Unknown cell edit kind.")
        result.append(CellEdit(edit.sheet, canonical_ref, edit.kind, edit.payload))
    return tuple(result)


def _validate_write_value(value: object, sheet: str, ref: str) -> None:
    if isinstance(value, (datetime, date, datetime_time)):
        raise ExcelLSPError(
            ErrorCode.INVALID_VALUE,
            f"Datetime writes are not supported in v0.1.0: {sheet}!{ref}.",
            hint="Write an Excel serial number or an ISO string instead.",
        )
    if value is None or isinstance(value, bool):
        return
    if isinstance(value, str):
        text_length = _utf16_code_units(value, sheet=sheet, ref=ref, kind="String")
        if text_length > _MAX_CELL_TEXT_LENGTH:
            raise ExcelLSPError(
                ErrorCode.INVALID_VALUE,
                f"String write at {sheet}!{ref} exceeds Excel's cell-text limit.",
                hint=f"Use at most {_MAX_CELL_TEXT_LENGTH} characters.",
            )
        return
    if type(value) is int:
        try:
            finite = math.isfinite(float(value))
        except OverflowError:
            finite = False
        if not finite:
            raise ExcelLSPError(
                ErrorCode.INVALID_VALUE,
                f"Numeric write at {sheet}!{ref} exceeds Excel's finite range.",
            )
        return
    if type(value) is float and math.isfinite(value):
        return
    raise ExcelLSPError(
        ErrorCode.INVALID_VALUE,
        f"Unsupported cell value at {sheet}!{ref}.",
        hint="Use a finite number, string, boolean, or null.",
    )


def _utf16_code_units(text: str, *, sheet: str, ref: str, kind: str) -> int:
    try:
        return len(text.encode("utf-16-le")) // 2
    except UnicodeError as exc:
        raise ExcelLSPError(
            ErrorCode.INVALID_VALUE,
            f"{kind} write at {sheet}!{ref} contains invalid Unicode text.",
        ) from exc


def _patch_sheet_xml(
    data: bytes,
    sheet_name: str,
    edits: Sequence[CellEdit],
) -> tuple[bytes, tuple[PatchedCell, ...]]:
    root = parse_xml(data)
    sheet_data = child_by_local(root, "sheetData")
    if sheet_data is None:
        raise _corrupt(f"worksheet {sheet_name!r} has no sheetData element")
    rows, cells = _index_sheet_cells(sheet_data)
    original_bounds = _cell_bounds(cells)

    array_spans = _array_formula_spans(cells)
    for edit in edits:
        row, col = parse_cell_ref(edit.ref)
        for span in array_spans:
            if (
                span.row_min <= row <= span.row_max
                and span.col_min <= col <= span.col_max
                and (span.row_min != span.row_max or span.col_min != span.col_max)
            ):
                raise ExcelLSPError(
                    ErrorCode.ARRAY_FORMULA,
                    f"Cannot edit {sheet_name}!{edit.ref} inside a multi-cell array formula.",
                    hint=("Rewrite the whole array or use a dynamic-array formula in its anchor."),
                )

    requested_keys = {parse_cell_ref(edit.ref) for edit in edits}
    expanded: set[tuple[int, int]] = set()
    shared_indexes: set[str] = set()
    for key in requested_keys:
        cell = cells.get(key)
        formula = None if cell is None else child_by_local(cell, "f")
        if formula is not None and attr_by_local(formula, "t") == "shared":
            shared_index = attr_by_local(formula, "si")
            if shared_index is None:
                raise _corrupt("shared formula has no si identity")
            shared_indexes.add(shared_index)
    for shared_index in sorted(shared_indexes):
        expanded.update(_expand_shared_formula_group(cells, shared_index))

    for edit in sorted(edits, key=lambda item: parse_cell_ref(item.ref)):
        row_number, column_number = parse_cell_ref(edit.ref)
        cell = cells.get((row_number, column_number))
        if cell is None:
            row = rows.get(row_number)
            if row is None:
                row = _insert_row(sheet_data, rows, row_number)
            cell = _insert_cell(row, cells, row_number, column_number)
        _apply_edit(cell, edit)

    extends = _edits_extend_bounds(requested_keys, original_bounds)
    _update_dimension(root, cells, extends=extends)
    patched = tuple(
        PatchedCell(
            sheet=sheet_name,
            ref=make_cell_ref(row, col),
            requested=(row, col) in requested_keys,
        )
        for row, col in sorted(requested_keys | expanded)
    )
    return _serialize_xml(root, data), patched


def _index_sheet_cells(
    sheet_data: etree._Element,
) -> tuple[dict[int, etree._Element], dict[tuple[int, int], etree._Element]]:
    rows: dict[int, etree._Element] = {}
    cells: dict[tuple[int, int], etree._Element] = {}
    previous_row = 0
    for row in sheet_data:
        if not isinstance(row.tag, str) or local_name(row.tag) != "row":
            continue
        raw_row = attr_by_local(row, "r")
        if raw_row is None or not raw_row.isascii() or not raw_row.isdigit():
            raise _corrupt("worksheet row has no valid r coordinate")
        row_number = int(raw_row)
        if not 1 <= row_number <= 1_048_576 or row_number <= previous_row:
            raise _corrupt("worksheet rows are duplicated, unordered, or out of bounds")
        previous_row = row_number
        rows[row_number] = row
        previous_col = 0
        for cell in row:
            if not isinstance(cell.tag, str) or local_name(cell.tag) != "c":
                continue
            raw_ref = attr_by_local(cell, "r")
            if raw_ref is None:
                raise _corrupt("worksheet cell has no r coordinate")
            try:
                cell_row, cell_col = parse_cell_ref(raw_ref)
            except ValueError as exc:
                raise _corrupt(f"worksheet cell has an invalid coordinate: {raw_ref!r}") from exc
            if cell_row != row_number or cell_col <= previous_col:
                raise _corrupt("worksheet cells are duplicated, unordered, or on the wrong row")
            previous_col = cell_col
            cells[(cell_row, cell_col)] = cell
    return rows, cells


def _array_formula_spans(
    cells: Mapping[tuple[int, int], etree._Element],
) -> tuple[Rect, ...]:
    spans = []
    for cell in cells.values():
        formula = child_by_local(cell, "f")
        if formula is None or attr_by_local(formula, "t") != "array":
            continue
        raw_ref = attr_by_local(formula, "ref") or attr_by_local(cell, "r")
        if raw_ref is None:
            raise _corrupt("array formula has no range")
        try:
            spans.append(parse_rect(raw_ref))
        except (UnicodeError, ValueError) as exc:
            raise _corrupt("array formula has an invalid range") from exc
    return tuple(spans)


def _expand_shared_formula_group(
    cells: Mapping[tuple[int, int], etree._Element],
    shared_index: str,
) -> set[tuple[int, int]]:
    members: list[tuple[tuple[int, int], etree._Element]] = []
    masters: list[tuple[tuple[int, int], etree._Element]] = []
    for key, cell in cells.items():
        formula = child_by_local(cell, "f")
        if formula is None:
            continue
        if attr_by_local(formula, "t") != "shared" or attr_by_local(formula, "si") != shared_index:
            continue
        members.append((key, formula))
        if formula.text and attr_by_local(formula, "ref"):
            masters.append((key, formula))
    if len(masters) != 1:
        raise _corrupt(f"shared formula group {shared_index!r} has no unique master")
    master_key, master = masters[0]
    raw_range = attr_by_local(master, "ref")
    assert raw_range is not None
    try:
        group_rect = parse_rect(raw_range)
    except (UnicodeError, ValueError) as exc:
        raise _corrupt(f"shared formula group {shared_index!r} has an invalid range") from exc
    expected_count = (group_rect.row_max - group_rect.row_min + 1) * (
        group_rect.col_max - group_rect.col_min + 1
    )
    if len(members) != expected_count:
        raise _corrupt(f"shared formula group {shared_index!r} is incomplete")
    master_text = master.text
    assert master_text is not None
    changed: set[tuple[int, int]] = set()
    for key, formula in members:
        row, col = key
        if not (
            group_rect.row_min <= row <= group_rect.row_max
            and group_rect.col_min <= col <= group_rect.col_max
        ):
            raise _corrupt(f"shared formula group {shared_index!r} escapes its declared range")
        try:
            translated = translate_a1_formula(
                f"={master_text}",
                origin=CellRef(*master_key),
                target=CellRef(row, col),
            )
        except ValueError as exc:
            raise _corrupt(f"shared formula group {shared_index!r} cannot be expanded") from exc
        formula.text = translated[1:]
        for attribute in tuple(formula.attrib):
            if local_name(attribute) in {"t", "si", "ref"}:
                del formula.attrib[attribute]
        changed.add(key)
    return changed


def _apply_edit(cell: etree._Element, edit: CellEdit) -> None:
    for child in tuple(cell):
        if isinstance(child.tag, str) and local_name(child.tag) in {"f", "v", "is"}:
            cell.remove(child)
    for attribute in tuple(cell.attrib):
        if local_name(attribute) == "t":
            del cell.attrib[attribute]

    namespace = _namespace(cell)
    if edit.kind is CellEditKind.FORMULA:
        formula = etree.Element(f"{{{namespace}}}f" if namespace else "f")
        try:
            formula.text = str(edit.payload)[1:]
        except ValueError as exc:
            raise ExcelLSPError(
                ErrorCode.INVALID_VALUE,
                f"Formula write at {edit.sheet}!{edit.ref} contains invalid XML characters.",
            ) from exc
        _insert_cell_content(cell, formula)
        return

    value = edit.payload
    if value is None:
        return
    if isinstance(value, bool):
        cell.set("t", "b")
        value_element = etree.Element(f"{{{namespace}}}v" if namespace else "v")
        value_element.text = "1" if value else "0"
        _insert_cell_content(cell, value_element)
        return
    if type(value) in {int, float}:
        value_element = etree.Element(f"{{{namespace}}}v" if namespace else "v")
        value_element.text = (
            str(int(value)) if type(value) is float and value.is_integer() else str(value)
        )
        _insert_cell_content(cell, value_element)
        return
    assert isinstance(value, str)
    cell.set("t", "inlineStr")
    inline = etree.Element(f"{{{namespace}}}is" if namespace else "is")
    text = etree.SubElement(inline, f"{{{namespace}}}t" if namespace else "t")
    text.set(_XML_SPACE, "preserve")
    try:
        text.text = value
    except ValueError as exc:
        raise ExcelLSPError(
            ErrorCode.INVALID_VALUE,
            f"String write at {edit.sheet}!{edit.ref} contains invalid XML characters.",
        ) from exc
    _insert_cell_content(cell, inline)


def _insert_cell_content(cell: etree._Element, content: etree._Element) -> None:
    """Insert CT_Cell value content before its optional trailing extension list."""
    for index, child in enumerate(cell):
        if isinstance(child.tag, str) and local_name(child.tag) == "extLst":
            cell.insert(index, content)
            return
    cell.append(content)


def _insert_row(
    sheet_data: etree._Element,
    rows: dict[int, etree._Element],
    row_number: int,
) -> etree._Element:
    namespace = _namespace(sheet_data)
    row = etree.Element(f"{{{namespace}}}row" if namespace else "row", r=str(row_number))
    insert_at = len(sheet_data)
    for index, existing in enumerate(sheet_data):
        if not isinstance(existing.tag, str) or local_name(existing.tag) != "row":
            continue
        existing_number = int(attr_by_local(existing, "r") or "0")
        if existing_number > row_number:
            insert_at = index
            break
    sheet_data.insert(insert_at, row)
    rows[row_number] = row
    return row


def _insert_cell(
    row: etree._Element,
    cells: dict[tuple[int, int], etree._Element],
    row_number: int,
    column_number: int,
) -> etree._Element:
    namespace = _namespace(row)
    ref = make_cell_ref(row_number, column_number)
    cell = etree.Element(f"{{{namespace}}}c" if namespace else "c", r=ref)
    insert_at = len(row)
    for index, existing in enumerate(row):
        if not isinstance(existing.tag, str) or local_name(existing.tag) != "c":
            continue
        raw_ref = attr_by_local(existing, "r")
        if raw_ref is None:
            raise _corrupt("worksheet cell has no r coordinate")
        _, existing_col = parse_cell_ref(raw_ref)
        if existing_col > column_number:
            insert_at = index
            break
    row.insert(insert_at, cell)
    cells[(row_number, column_number)] = cell
    return cell


def _cell_bounds(
    cells: Mapping[tuple[int, int], etree._Element],
) -> tuple[int, int, int, int] | None:
    if not cells:
        return None
    return (
        min(row for row, _ in cells),
        max(row for row, _ in cells),
        min(col for _, col in cells),
        max(col for _, col in cells),
    )


def _edits_extend_bounds(
    edited: Iterable[tuple[int, int]],
    original: tuple[int, int, int, int] | None,
) -> bool:
    if original is None:
        return True
    row_min, row_max, col_min, col_max = original
    return any(
        row < row_min or row > row_max or col < col_min or col > col_max for row, col in edited
    )


def _update_dimension(
    root: etree._Element,
    cells: Mapping[tuple[int, int], etree._Element],
    *,
    extends: bool,
) -> None:
    dimensions = [
        child
        for child in root
        if isinstance(child.tag, str) and local_name(child.tag) == "dimension"
    ]
    for dimension in dimensions:
        root.remove(dimension)
    if not extends:
        return
    bounds = _cell_bounds(cells)
    if bounds is None:
        return
    row_min, row_max, col_min, col_max = bounds
    first = make_cell_ref(row_min, col_min)
    last = make_cell_ref(row_max, col_max)
    namespace = _namespace(root)
    dimension = etree.Element(f"{{{namespace}}}dimension" if namespace else "dimension")
    dimension.set("ref", first if first == last else f"{first}:{last}")
    insert_at = 1 if len(root) and local_name(root[0].tag) == "sheetPr" else 0
    root.insert(insert_at, dimension)


def _ensure_full_calculation(data: bytes) -> bytes:
    root = parse_xml(data)
    calc_properties = [
        child for child in root if isinstance(child.tag, str) and local_name(child.tag) == "calcPr"
    ]
    if len(calc_properties) > 1:
        raise _corrupt("workbook has duplicate calcPr elements")
    if calc_properties:
        calc = calc_properties[0]
    else:
        namespace = _namespace(root)
        calc = etree.Element(f"{{{namespace}}}calcPr" if namespace else "calcPr")
        insert_at = len(root)
        for index, child in enumerate(root):
            if isinstance(child.tag, str) and local_name(child.tag) in _CALC_PR_SUCCESSORS:
                insert_at = index
                break
        root.insert(insert_at, calc)
    if attr_by_local(calc, "fullCalcOnLoad") == "1":
        return data
    for attribute in tuple(calc.attrib):
        if local_name(attribute) == "fullCalcOnLoad":
            del calc.attrib[attribute]
    calc.set("fullCalcOnLoad", "1")
    return _serialize_xml(root, data)


def _patch_calc_chain_references(
    source: ZipFile,
    infos: Mapping[str, ZipInfo],
    modified_payloads: dict[str, bytes],
    deleted_parts: set[str],
) -> None:
    relationship_info = infos.get(_WORKBOOK_RELS_PART)
    content_types_info = infos.get(_CONTENT_TYPES_PART)
    if relationship_info is None or content_types_info is None:
        raise _corrupt("workbook relationship or content-type catalog is missing")

    rel_original = _read_member(source, relationship_info)
    rel_root = parse_xml(rel_original)
    rel_changed = False
    for relationship in tuple(rel_root):
        if not isinstance(relationship.tag, str) or local_name(relationship.tag) != "Relationship":
            continue
        relationship_type = (attr_by_local(relationship, "Type") or "").casefold()
        target = attr_by_local(relationship, "Target") or ""
        target_part = _resolve_relationship_target(_WORKBOOK_PART, target)
        if (
            relationship_type.endswith("/calcchain")
            or target_part.casefold() == _CALC_CHAIN_PART.casefold()
        ):
            rel_root.remove(relationship)
            rel_changed = True
    if rel_changed:
        modified_payloads[_WORKBOOK_RELS_PART] = _serialize_xml(rel_root, rel_original)

    types_original = _read_member(source, content_types_info)
    types_root = parse_xml(types_original)
    types_changed = False
    for override in tuple(types_root):
        if not isinstance(override.tag, str) or local_name(override.tag) != "Override":
            continue
        part_name = _normalize_part_name(attr_by_local(override, "PartName") or "")
        if part_name.casefold() == _CALC_CHAIN_PART.casefold():
            types_root.remove(override)
            types_changed = True
    if types_changed:
        modified_payloads[_CONTENT_TYPES_PART] = _serialize_xml(types_root, types_original)

    if deleted_parts and (not rel_changed or not types_changed):
        # Real workbooks occasionally omit one catalog entry.  Removing every
        # representation that exists is safe; absence is not a dangling link.
        return


def _check_excel_lockfile(workbook: Path) -> None:
    lockfile = workbook.with_name(f"~${workbook.name}")
    if lockfile.exists():
        raise ExcelLSPError(
            ErrorCode.OPEN_IN_EXCEL,
            f"Workbook appears to be open in Excel: {workbook.name}.",
            hint=(
                "Close it in Excel first; remove a stale lockfile only after "
                "confirming Excel is closed."
            ),
        )


def _index_archive(source: ZipFile) -> dict[str, ZipInfo]:
    infos: dict[str, ZipInfo] = {}
    for info in source.infolist():
        normalized = _normalize_part_name(info.filename)
        if not info.is_dir() and normalized in infos:
            raise _corrupt(f"duplicate package member: {normalized}")
        if info.flag_bits & 0x1:
            raise ExcelLSPError(
                ErrorCode.ENCRYPTED,
                "Encrypted workbook parts cannot be edited.",
            )
        if not info.is_dir():
            infos[normalized] = info
    return infos


def _write_replacement_archive(
    workbook: Path,
    source: ZipFile,
    modified_payloads: Mapping[str, bytes],
    deleted_parts: set[str],
) -> Path:
    descriptor, raw_path = tempfile.mkstemp(
        prefix=f".{workbook.name}.",
        suffix=workbook.suffix,
        dir=workbook.parent,
    )
    os.close(descriptor)
    temporary = Path(raw_path)
    try:
        with ZipFile(temporary, "w", allowZip64=True) as target:
            target.comment = source.comment
            for info in source.infolist():
                normalized = _normalize_part_name(info.filename)
                if normalized in deleted_parts:
                    continue
                payload = (
                    b""
                    if info.is_dir()
                    else modified_payloads.get(normalized, _read_member(source, info))
                )
                target.writestr(info, payload, compress_type=info.compress_type)
        with temporary.open("rb+") as stream:
            os.fsync(stream.fileno())
        return temporary
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise


def _validate_replacement(path: Path, sheet_names: Iterable[str]) -> str:
    """Reparse every touched sheet before the atomic replacement boundary."""
    with OOXMLParser(path) as parser:
        descriptors = {descriptor.name: descriptor for descriptor in parser.metadata.sheets}
        for sheet_name in sorted(set(sheet_names)):
            descriptor = descriptors.get(sheet_name)
            if descriptor is None:
                raise _corrupt(f"patched worksheet disappeared: {sheet_name!r}")
            parser.parse_sheet(descriptor, lambda _cell: None)
        return parser.hashes.whole_file


def _atomic_replace_with_retry(
    source: Path,
    destination: Path,
    *,
    expected_workbook_hash: str,
) -> None:
    for attempt in range(2):
        _check_excel_lockfile(destination)
        if _hash_file(destination) != expected_workbook_hash:
            raise ExcelLSPError(
                ErrorCode.CONFLICT,
                "Workbook changed before the prepared edit could be installed.",
                hint="Run refresh, review the changed workbook, and retry the edit.",
            )
        try:
            os.replace(source, destination)
            return
        except PermissionError as exc:
            if attempt == 0:
                time.sleep(_REPLACE_RETRY_DELAY_SECONDS)
                continue
            raise ExcelLSPError(
                ErrorCode.LOCKED,
                f"Workbook could not be replaced: {destination.name}.",
                hint="Close the application holding the file and retry.",
            ) from exc
    raise AssertionError("unreachable replacement retry state")


def _read_member(source: ZipFile, info: ZipInfo) -> bytes:
    try:
        return source.read(info)
    except RuntimeError as exc:
        if "password" in str(exc).casefold() or info.flag_bits & 0x1:
            raise ExcelLSPError(ErrorCode.ENCRYPTED, "Workbook part is encrypted.") from exc
        raise


def _resolve_relationship_target(source_part: str, target: str) -> str:
    if not target:
        return ""
    if target.startswith("/"):
        return _normalize_part_name(target)
    return _normalize_part_name(posixpath.join(posixpath.dirname(source_part), target))


def _normalize_part_name(part_name: str) -> str:
    candidate = part_name.replace("\\", "/").lstrip("/")
    normalized = posixpath.normpath(candidate)
    if normalized in {"", "."} or normalized == ".." or normalized.startswith("../"):
        return candidate
    return normalized


def _namespace(element: etree._Element) -> str:
    if isinstance(element.tag, str) and element.tag.startswith("{"):
        return element.tag[1:].split("}", 1)[0]
    return ""


def _serialize_xml(root: etree._Element, original: bytes) -> bytes:
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=original.lstrip().startswith(b"<?xml"),
    )


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def _corrupt(message: str) -> ExcelLSPError:
    return ExcelLSPError(
        ErrorCode.CORRUPT,
        message,
        hint="Open and resave the workbook in Excel, then retry.",
    )


__all__ = ["patch_workbook"]
