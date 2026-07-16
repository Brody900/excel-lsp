"""Build the deterministic Excel workbooks used by the test suite.

openpyxl is intentionally limited to authoring new fixture workbooks. The
post-processing helpers below patch raw OOXML so formula caches and shared
formula groups match files saved by Excel, then repack every archive with a
stable member order and timestamp.
"""

from __future__ import annotations

import math
import posixpath
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Literal
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import lxml.etree as etree
from openpyxl import Workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

GENERATED_DIR = Path(__file__).resolve().parent / "generated"

_DOCUMENT_TIMESTAMP = datetime(2000, 1, 1, 0, 0, 0)
_DOCUMENT_TIMESTAMP_TEXT = "2000-01-01T00:00:00Z"
_ZIP_TIMESTAMP = (1980, 1, 1, 0, 0, 0)
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_DCTERMS_NS = "http://purl.org/dc/terms/"

CachedValueKind = Literal["number", "string", "error", "boolean"]


@dataclass(frozen=True, slots=True)
class CachedValue:
    """A formula result known by the fixture author, not a general evaluator."""

    value: int | float | str | bool
    kind: CachedValueKind = "number"


@dataclass(frozen=True, slots=True)
class SharedFormulaGroup:
    """One OOXML shared-formula rectangle and its top-left master cell."""

    range_ref: str
    master_ref: str
    shared_index: int


def _read_archive(path: Path) -> dict[str, bytes]:
    with ZipFile(path, "r") as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _write_deterministic_archive(path: Path, members: Mapping[str, bytes]) -> None:
    stable_members = dict(members)
    core_properties = stable_members.get("docProps/core.xml")
    if core_properties is not None:
        core = etree.fromstring(core_properties)
        for property_name in ("created", "modified"):
            element = core.find(f"{{{_DCTERMS_NS}}}{property_name}")
            if element is not None:
                element.text = _DOCUMENT_TIMESTAMP_TEXT
        stable_members["docProps/core.xml"] = etree.tostring(core, encoding="utf-8")

    temporary = path.with_name(f".{path.name}.tmp")
    temporary.unlink(missing_ok=True)
    try:
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
            for name in sorted(stable_members):
                info = ZipInfo(filename=name, date_time=_ZIP_TIMESTAMP)
                info.compress_type = ZIP_DEFLATED
                info.create_system = 0
                info.external_attr = 0
                info.internal_attr = 0
                info.extra = b""
                info.comment = b""
                archive.writestr(
                    info,
                    stable_members[name],
                    compress_type=ZIP_DEFLATED,
                    compresslevel=9,
                )
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def repack_deterministic(path: Path) -> None:
    """Sort ZIP members and apply the earliest valid ZIP timestamp to each."""
    _write_deterministic_archive(path, _read_archive(path))


def _worksheet_part(members: Mapping[str, bytes], sheet_name: str) -> str:
    workbook = etree.fromstring(members["xl/workbook.xml"])
    sheet = next(
        (
            candidate
            for candidate in workbook.findall(f".//{{{_MAIN_NS}}}sheet")
            if candidate.get("name") == sheet_name
        ),
        None,
    )
    if sheet is None:
        raise ValueError(f"worksheet not found: {sheet_name}")

    relationship_id = sheet.get(f"{{{_OFFICE_REL_NS}}}id")
    relationships = etree.fromstring(members["xl/_rels/workbook.xml.rels"])
    relationship = next(
        (
            candidate
            for candidate in relationships.findall(f"{{{_PACKAGE_REL_NS}}}Relationship")
            if candidate.get("Id") == relationship_id
        ),
        None,
    )
    if relationship is None or not relationship.get("Target"):
        raise ValueError(f"worksheet relationship is missing: {sheet_name}")

    target = relationship.get("Target")
    assert target is not None
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def _serialized_cached_value(cached: CachedValue) -> tuple[str | None, str]:
    if cached.kind == "number":
        if isinstance(cached.value, bool) or not isinstance(cached.value, (int, float)):
            raise TypeError("numeric cached values must be int or float, not bool")
        if isinstance(cached.value, float) and not math.isfinite(cached.value):
            raise ValueError("numeric cached values must be finite")
        return None, str(cached.value)
    if cached.kind == "boolean":
        if not isinstance(cached.value, bool):
            raise TypeError("boolean cached values must be bool")
        return "b", "1" if cached.value else "0"
    if not isinstance(cached.value, str):
        raise TypeError(f"{cached.kind} cached values must be strings")
    return ("str" if cached.kind == "string" else "e"), cached.value


def inject_cached_values(
    workbook_path: Path,
    sheet_name: str,
    cached_values: Mapping[str, CachedValue],
) -> None:
    """Inject fixture-author-computed formula results into one worksheet part."""
    members = _read_archive(workbook_path)
    part = _worksheet_part(members, sheet_name)
    worksheet = etree.fromstring(members[part])
    cells = {
        cell.get("r"): cell for cell in worksheet.findall(f".//{{{_MAIN_NS}}}c") if cell.get("r")
    }

    for ref, cached in cached_values.items():
        cell = cells.get(ref)
        if cell is None:
            raise ValueError(f"cannot cache missing cell: {sheet_name}!{ref}")
        formula = cell.find(f"{{{_MAIN_NS}}}f")
        if formula is None:
            raise ValueError(f"cannot cache non-formula cell: {sheet_name}!{ref}")

        cell_type, serialized = _serialized_cached_value(cached)
        if cell_type is None:
            cell.attrib.pop("t", None)
        else:
            cell.set("t", cell_type)

        for existing in cell.findall(f"{{{_MAIN_NS}}}v"):
            cell.remove(existing)
        value = etree.Element(f"{{{_MAIN_NS}}}v")
        value.text = serialized
        cell.insert(cell.index(formula) + 1, value)

    members[part] = etree.tostring(worksheet, encoding="utf-8")
    _write_deterministic_archive(workbook_path, members)


def _refs_in_range(range_ref: str) -> tuple[str, ...]:
    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise ValueError(f"shared formula range must be bounded: {range_ref}")
    return tuple(
        f"{get_column_letter(column)}{row}"
        for row in range(min_row, max_row + 1)
        for column in range(min_col, max_col + 1)
    )


def convert_shared_formula_groups(
    workbook_path: Path,
    sheet_name: str,
    groups: Sequence[SharedFormulaGroup],
) -> None:
    """Convert full formulas to OOXML shared masters and si-only followers."""
    members = _read_archive(workbook_path)
    part = _worksheet_part(members, sheet_name)
    worksheet = etree.fromstring(members[part])
    cells = {
        cell.get("r"): cell for cell in worksheet.findall(f".//{{{_MAIN_NS}}}c") if cell.get("r")
    }

    seen_indexes: set[int] = set()
    for group in groups:
        if group.shared_index in seen_indexes:
            raise ValueError(f"duplicate shared formula index: {group.shared_index}")
        seen_indexes.add(group.shared_index)

        refs = _refs_in_range(group.range_ref)
        if not refs or refs[0] != group.master_ref:
            raise ValueError("shared formula master must be the top-left cell of its range")

        formulas = {}
        for ref in refs:
            cell = cells.get(ref)
            formula = None if cell is None else cell.find(f"{{{_MAIN_NS}}}f")
            if formula is None or not formula.text:
                raise ValueError(f"shared formula source is missing: {sheet_name}!{ref}")
            formulas[ref] = formula

        master = formulas[group.master_ref]
        master.attrib.clear()
        master.set("t", "shared")
        master.set("ref", group.range_ref)
        master.set("si", str(group.shared_index))

        for ref in refs[1:]:
            follower = formulas[ref]
            follower.attrib.clear()
            follower.set("t", "shared")
            follower.set("si", str(group.shared_index))
            follower.text = None

    members[part] = etree.tostring(worksheet, encoding="utf-8")
    _write_deterministic_archive(workbook_path, members)


def _new_workbook(sheet_name: str) -> tuple[Workbook, Worksheet]:
    workbook = Workbook()
    worksheet = workbook.active
    if not isinstance(worksheet, Worksheet):
        raise RuntimeError("new workbook did not create a worksheet")
    worksheet.title = sheet_name
    workbook.properties.creator = "Excel LSP fixture generator"
    workbook.properties.lastModifiedBy = "Excel LSP fixture generator"
    workbook.properties.created = _DOCUMENT_TIMESTAMP
    workbook.properties.modified = _DOCUMENT_TIMESTAMP
    workbook.properties.title = f"Excel LSP {sheet_name} fixture"
    return workbook, worksheet


def _table(name: str, ref: str) -> Table:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    return table


def _generate_f01(output_dir: Path) -> Path:
    path = output_dir / "basic_single_table.xlsx"
    workbook, worksheet = _new_workbook("Sales")
    worksheet.append(("Item", "Quantity", "UnitPrice", "LineTotal"))

    rows = (
        ("Widget", 2, 3.5),
        ("Gadget", 5, 1.25),
        ("Bracket", 3, 4.0),
        ("Cable", 7, 2.0),
        ("Adapter", 4, 6.75),
    )
    caches: dict[str, CachedValue] = {}
    for row_number, (item, quantity, unit_price) in enumerate(rows, start=2):
        worksheet.append((item, quantity, unit_price, f"=B{row_number}*C{row_number}"))
        caches[f"D{row_number}"] = CachedValue(quantity * unit_price)

    worksheet.add_table(_table("SalesTable", "A1:D6"))
    workbook.save(path)
    workbook.close()
    inject_cached_values(path, "Sales", caches)
    repack_deterministic(path)
    return path


def _generate_f07(output_dir: Path) -> Path:
    path = output_dir / "formula_blocks.xlsx"
    workbook, worksheet = _new_workbook("FormulaBlocks")
    worksheet.append(("Factor", "Multiplier", "Product"))

    caches: dict[str, CachedValue] = {}
    for row_number in range(2, 22):
        factor = row_number - 1
        multiplier = (row_number % 5) + 2
        formula = (
            f"=A{row_number}+B{row_number}" if row_number == 12 else f"=A{row_number}*B{row_number}"
        )
        worksheet.append((factor, multiplier, formula))
        expected = factor + multiplier if row_number == 12 else factor * multiplier
        caches[f"C{row_number}"] = CachedValue(expected)

    worksheet.add_table(_table("FormulaBlocksTable", "A1:C21"))
    worksheet["E1"] = "Read-only metadata probe"
    worksheet.merge_cells("E1:F1")
    workbook.save(path)
    workbook.close()

    inject_cached_values(path, "FormulaBlocks", caches)
    convert_shared_formula_groups(
        path,
        "FormulaBlocks",
        (
            SharedFormulaGroup("C2:C11", "C2", 0),
            SharedFormulaGroup("C13:C21", "C13", 1),
        ),
    )
    repack_deterministic(path)
    return path


def generate_all(output_dir: Path = GENERATED_DIR) -> dict[str, Path]:
    """Generate the Phase 1 fixture subset and return paths keyed by fixture ID."""
    output_dir.mkdir(parents=True, exist_ok=True)
    return {
        "F01": _generate_f01(output_dir),
        "F07": _generate_f07(output_dir),
    }


def main() -> None:
    """Generate fixtures into ``tests/fixtures/generated``."""
    for fixture_id, path in generate_all().items():
        print(f"{fixture_id}: {path}")


if __name__ == "__main__":
    main()
