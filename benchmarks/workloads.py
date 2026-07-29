"""Deterministic benchmark workload envelopes built from canonical fixtures."""

from __future__ import annotations

import posixpath
import re
from collections.abc import Mapping
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import lxml.etree as etree

ARCHIVE_ROWS = 1_000
WORKLOAD_SUBDIR = Path("tests/fixtures/generated/benchmarks")
FIXTURES = (
    "cross_sheet_model.xlsx",
    "formula_blocks.xlsx",
    "errors.xlsx",
    "mixed_types.xlsx",
)
_HEADERS = (
    "RecordID",
    "Region",
    "Channel",
    "Product",
    "Units",
    "Price",
    "Discount",
    "Status",
)
_ZIP_TIMESTAMP = (1980, 1, 1, 0, 0, 0)
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_WORKSHEET_REL = f"{_OFFICE_REL_NS}/worksheet"
_WORKSHEET_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)
_COLUMNS = "ABCDEFGH"


def _archive_row(row_number: int) -> tuple[object, ...]:
    return (
        row_number,
        f"Region-{row_number % 12:02d}",
        f"Channel-{row_number % 4}",
        f"SKU-{row_number % 80:03d}",
        1 + row_number % 25,
        10.25 + row_number % 90,
        (row_number % 7) / 100,
        "closed",
    )


def _read_archive(path: Path) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _write_archive(path: Path, members: Mapping[str, bytes]) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.unlink(missing_ok=True)
    try:
        with ZipFile(temporary, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
            for name in sorted(members):
                info = ZipInfo(filename=name, date_time=_ZIP_TIMESTAMP)
                info.compress_type = ZIP_DEFLATED
                info.create_system = 0
                info.external_attr = 0
                info.internal_attr = 0
                info.extra = b""
                info.comment = b""
                archive.writestr(info, members[name], compress_type=ZIP_DEFLATED, compresslevel=9)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def _append_cell(row: etree._Element, row_number: int, column: int, value: object) -> None:
    cell = etree.SubElement(row, f"{{{_MAIN_NS}}}c", r=f"{_COLUMNS[column - 1]}{row_number}")
    if isinstance(value, str):
        cell.set("t", "inlineStr")
        inline = etree.SubElement(cell, f"{{{_MAIN_NS}}}is")
        etree.SubElement(inline, f"{{{_MAIN_NS}}}t").text = value
    else:
        etree.SubElement(cell, f"{{{_MAIN_NS}}}v").text = str(value)


def _append_archive_rows(sheet_data: etree._Element, *, start_row: int) -> None:
    header = etree.SubElement(sheet_data, f"{{{_MAIN_NS}}}row", r=str(start_row))
    for column, value in enumerate(_HEADERS, start=1):
        _append_cell(header, start_row, column, value)
    for offset in range(1, ARCHIVE_ROWS + 1):
        row_number = start_row + offset
        row = etree.SubElement(sheet_data, f"{{{_MAIN_NS}}}row", r=str(row_number))
        for column, value in enumerate(_archive_row(offset), start=1):
            _append_cell(row, row_number, column, value)


def _new_archive_worksheet() -> bytes:
    worksheet = etree.Element(f"{{{_MAIN_NS}}}worksheet", nsmap={None: _MAIN_NS})
    etree.SubElement(
        worksheet,
        f"{{{_MAIN_NS}}}dimension",
        ref=f"A1:H{ARCHIVE_ROWS + 1}",
    )
    views = etree.SubElement(worksheet, f"{{{_MAIN_NS}}}sheetViews")
    etree.SubElement(views, f"{{{_MAIN_NS}}}sheetView", workbookViewId="0")
    etree.SubElement(
        worksheet, f"{{{_MAIN_NS}}}sheetFormatPr", baseColWidth="10", defaultRowHeight="15"
    )
    sheet_data = etree.SubElement(worksheet, f"{{{_MAIN_NS}}}sheetData")
    _append_archive_rows(sheet_data, start_row=1)
    etree.SubElement(
        worksheet,
        f"{{{_MAIN_NS}}}pageMargins",
        left="0.75",
        right="0.75",
        top="1",
        bottom="1",
        header="0.5",
        footer="0.5",
    )
    return etree.tostring(worksheet, encoding="utf-8", xml_declaration=True, standalone=True)


def _next_relationship_id(relationships: etree._Element) -> str:
    used = {relationship.get("Id", "") for relationship in relationships}
    numeric = [int(match.group(1)) for value in used if (match := re.fullmatch(r"rId(\d+)", value))]
    candidate = max(numeric, default=0) + 1
    while f"rId{candidate}" in used:
        candidate += 1
    return f"rId{candidate}"


def _worksheet_part(members: Mapping[str, bytes], sheet_name: str) -> str:
    workbook = etree.fromstring(members["xl/workbook.xml"])
    sheet = next(
        candidate
        for candidate in workbook.findall(f".//{{{_MAIN_NS}}}sheet")
        if candidate.get("name") == sheet_name
    )
    relationship_id = sheet.get(f"{{{_OFFICE_REL_NS}}}id")
    relationships = etree.fromstring(members["xl/_rels/workbook.xml.rels"])
    relationship = next(
        candidate
        for candidate in relationships.findall(f"{{{_PACKAGE_REL_NS}}}Relationship")
        if candidate.get("Id") == relationship_id
    )
    target = relationship.get("Target")
    if target is None:
        raise ValueError(f"worksheet relationship has no target: {sheet_name}")
    normalized = target.lstrip("/")
    if normalized.startswith("xl/"):
        return posixpath.normpath(normalized)
    return posixpath.normpath(posixpath.join("xl", normalized))


def _add_summary_archive(members: dict[str, bytes]) -> None:
    part = _worksheet_part(members, "Summary")
    worksheet = etree.fromstring(members[part])
    sheet_data = worksheet.find(f"{{{_MAIN_NS}}}sheetData")
    dimension = worksheet.find(f"{{{_MAIN_NS}}}dimension")
    if sheet_data is None or dimension is None:
        raise ValueError("Summary worksheet is missing required OOXML elements")
    _append_archive_rows(sheet_data, start_row=1_000)
    dimension.set("ref", "A1:H2000")
    members[part] = etree.tostring(worksheet, encoding="utf-8", xml_declaration=True)


def _add_archive_sheet(members: dict[str, bytes]) -> None:
    workbook = etree.fromstring(members["xl/workbook.xml"])
    sheets = workbook.find(f"{{{_MAIN_NS}}}sheets")
    if sheets is None:
        raise ValueError("workbook has no sheets collection")
    sheet_id = max(int(sheet.get("sheetId", "0")) for sheet in sheets) + 1

    relationships = etree.fromstring(members["xl/_rels/workbook.xml.rels"])
    relationship_id = _next_relationship_id(relationships)
    part_numbers = [
        int(match.group(1))
        for relationship in relationships.findall(f"{{{_PACKAGE_REL_NS}}}Relationship")
        if relationship.get("Type") == _WORKSHEET_REL
        and (match := re.search(r"sheet(\d+)\.xml$", relationship.get("Target", "")))
    ]
    part_number = max(part_numbers, default=0) + 1
    target = f"worksheets/sheet{part_number}.xml"
    part = f"xl/{target}"

    etree.SubElement(
        sheets,
        f"{{{_MAIN_NS}}}sheet",
        name="BenchmarkArchive",
        sheetId=str(sheet_id),
        attrib={f"{{{_OFFICE_REL_NS}}}id": relationship_id},
    )
    etree.SubElement(
        relationships,
        f"{{{_PACKAGE_REL_NS}}}Relationship",
        Id=relationship_id,
        Type=_WORKSHEET_REL,
        Target=target,
    )

    content_types = etree.fromstring(members["[Content_Types].xml"])
    etree.SubElement(
        content_types,
        f"{{{_CONTENT_TYPES_NS}}}Override",
        PartName=f"/{part}",
        ContentType=_WORKSHEET_CONTENT_TYPE,
    )

    members["xl/workbook.xml"] = etree.tostring(workbook, encoding="utf-8", xml_declaration=True)
    members["xl/_rels/workbook.xml.rels"] = etree.tostring(
        relationships, encoding="utf-8", xml_declaration=True
    )
    members["[Content_Types].xml"] = etree.tostring(
        content_types, encoding="utf-8", xml_declaration=True
    )
    members[part] = _new_archive_worksheet()


def build_workload(source: Path, destination: Path) -> Path:
    """Add disclosed history with only the documented OOXML member changes."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    members = _read_archive(source)
    if source.name == "cross_sheet_model.xlsx":
        _add_summary_archive(members)
    _add_archive_sheet(members)
    _write_archive(destination, members)
    return destination


def build_workloads(
    root: Path,
    *,
    output_dir: Path | None = None,
    force: bool = False,
) -> dict[str, Path]:
    """Build every benchmark workbook from the generated canonical fixtures."""
    source_dir = root / "tests" / "fixtures" / "generated"
    destination_dir = output_dir or root / WORKLOAD_SUBDIR
    result: dict[str, Path] = {}
    for fixture in FIXTURES:
        source = source_dir / fixture
        if not source.is_file():
            raise FileNotFoundError(
                f"missing generated fixture {source}; run tests/fixtures/generate.py first"
            )
        destination = destination_dir / fixture
        if (
            force
            or not destination.is_file()
            or destination.stat().st_mtime_ns < source.stat().st_mtime_ns
        ):
            build_workload(source, destination)
        result[fixture] = destination
    return result


__all__ = ["ARCHIVE_ROWS", "FIXTURES", "WORKLOAD_SUBDIR", "build_workload", "build_workloads"]
