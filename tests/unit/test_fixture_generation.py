from __future__ import annotations

import hashlib
import posixpath
import runpy
from collections import Counter
from collections.abc import Callable
from pathlib import Path
from typing import Any, cast
from zipfile import ZipFile

import lxml.etree as etree
import pytest
from openpyxl.utils.cell import get_column_letter

from excel_lsp.core.parse import OOXMLParser

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
FIXED_ZIP_TIMESTAMP = (1980, 1, 1, 0, 0, 0)
EXPECTED_FIXTURE_IDS = {
    "F01",
    "F02",
    "F03",
    "F04",
    "F05",
    "F07",
    "F08",
    "F09a",
    "F09b",
    "F10",
    "F11",
    "F12",
    "F13",
    "F14",
    "F15",
    "F18",
    "F19",
    "F20",
}
P1_SHA256 = {
    "F01": "8d57d9143edf78a66be6c33bcede3bcc7fba8ed1ac2d816391a4139a28a41270",
    "F07": "50015028edc75a4bab5cd13af9b4576f520d8c1f4cf0e3b223bab54c3476c871",
}
P3_SHA256 = {
    "F19": "db6ab279299ff1ed28d120fdc4ba6b057227867fc92e98caba50ca4ce761dcdf",
}
P4_SHA256 = {
    "F04": "9844b879673deae1455054aca03f20168e10121d61f23e3204235b8ce3574a0c",
    "F05": "1132fe9ddd6cca3d0a4e4af4a38e4a56d8b1ebbf1f9526e19c6dec053fc1d397",
    "F09a": "b18a88c6e1c92a25ef0c3de851bd5675278f03d04518f80db6f501416fbf1234",
    "F09b": "0998132a470b1258aa2b2c1a68162f5c7c16b59e56d8848b6110cbd2f8917675",
    "F15": "96496cf60b8990e4d97c857cfe00b4593c856e3bb687ec376cdc00fe830a788f",
}
P5_SHA256 = {
    "F08": "9e54d1d9f93947fcecfd6c90994beba91b50b63127f229897a132b39236876b4",
    "F10": "0418ee82e8d319842ede2e752350e23a1f4390a7834bd124b4ba288a50d533c6",
    "F11": "b800e80117a306bf39993ef82b328bae0d624a328051ed667ea63c577608438b",
    "F18": "5b77a7efc0d93e06355206bba1385dc14a3743f592a8864091913300e57bfc7e",
}
GenerateAll = Callable[[Path], dict[str, Path]]
generate_all = cast(
    GenerateAll,
    runpy.run_path(str(Path(__file__).parents[1] / "fixtures" / "generate.py"))["generate_all"],
)


@pytest.fixture(scope="module")
def generated_paths(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    """Generate the full corpus once for shape tests, including the 50k fixture."""
    return generate_all(tmp_path_factory.mktemp("fixture-corpus"))


def _xml_member(path: Path, member: str):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read(member))


def _worksheet_part(path: Path, sheet_name: str) -> str:
    workbook = _xml_member(path, "xl/workbook.xml")
    sheet = next(
        candidate
        for candidate in workbook.findall(f".//{{{MAIN_NS}}}sheet")
        if candidate.get("name") == sheet_name
    )
    relationship_id = sheet.get(f"{{{OFFICE_REL_NS}}}id")
    relationships = _xml_member(path, "xl/_rels/workbook.xml.rels")
    relationship = next(
        candidate
        for candidate in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
        if candidate.get("Id") == relationship_id
    )
    target = relationship.get("Target")
    assert target is not None
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def _worksheet(path: Path, sheet_name: str):
    return _xml_member(path, _worksheet_part(path, sheet_name))


def _cells(path: Path, sheet_name: str) -> dict[str, Any]:
    worksheet = _worksheet(path, sheet_name)
    return {
        cell.get("r"): cell for cell in worksheet.findall(f".//{{{MAIN_NS}}}c") if cell.get("r")
    }


def _rect_refs(min_col: int, min_row: int, max_col: int, max_row: int) -> set[str]:
    return {
        f"{get_column_letter(column)}{row}"
        for row in range(min_row, max_row + 1)
        for column in range(min_col, max_col + 1)
    }


def _cell_text(cell: Any) -> str:
    return "".join(cell.itertext())


def _formula_and_value(cell: Any) -> tuple[str | None, str | None]:
    formula = cell.find(f"{{{MAIN_NS}}}f")
    value = cell.find(f"{{{MAIN_NS}}}v")
    return (
        None if formula is None else formula.text,
        None if value is None else value.text,
    )


def _tables(path: Path) -> set[tuple[str | None, str | None]]:
    with ZipFile(path) as archive:
        parts = sorted(
            name
            for name in archive.namelist()
            if name.startswith("xl/tables/table") and name.endswith(".xml")
        )
        return {
            (root.get("name"), root.get("ref"))
            for root in (etree.fromstring(archive.read(part)) for part in parts)
        }


def test_generation_is_byte_identical_with_stable_zip_metadata(tmp_path: Path) -> None:
    first_paths = generate_all(tmp_path / "first")
    first_bytes = {fixture_id: path.read_bytes() for fixture_id, path in first_paths.items()}

    second_paths = generate_all(tmp_path / "second")

    assert second_paths.keys() == first_paths.keys() == EXPECTED_FIXTURE_IDS
    for fixture_id, path in second_paths.items():
        assert path.read_bytes() == first_bytes[fixture_id]
        with ZipFile(path) as archive:
            infos = archive.infolist()
        assert [info.filename for info in infos] == sorted(info.filename for info in infos)
        assert {info.date_time for info in infos} == {FIXED_ZIP_TIMESTAMP}

    assert {
        fixture_id: hashlib.sha256(first_bytes[fixture_id]).hexdigest() for fixture_id in P1_SHA256
    } == P1_SHA256
    assert {
        fixture_id: hashlib.sha256(first_bytes[fixture_id]).hexdigest() for fixture_id in P3_SHA256
    } == P3_SHA256
    assert {
        fixture_id: hashlib.sha256(first_bytes[fixture_id]).hexdigest() for fixture_id in P4_SHA256
    } == P4_SHA256
    assert {
        fixture_id: hashlib.sha256(first_bytes[fixture_id]).hexdigest() for fixture_id in P5_SHA256
    } == P5_SHA256


def test_f01_has_listobject_formulas_and_injected_caches(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F01"]
    cells = _cells(path, "Sales")
    expected_caches = {
        "D2": 7.0,
        "D3": 6.25,
        "D4": 12.0,
        "D5": 14.0,
        "D6": 27.0,
    }
    for ref, expected in expected_caches.items():
        cell = cells[ref]
        formula = cell.find(f"{{{MAIN_NS}}}f")
        value = cell.find(f"{{{MAIN_NS}}}v")
        assert formula is not None
        assert formula.text == f"B{ref[1:]}*C{ref[1:]}"
        assert value is not None
        assert float(value.text) == expected

    table = _xml_member(path, "xl/tables/table1.xml")
    assert table.get("name") == "SalesTable"
    assert table.get("displayName") == "SalesTable"
    assert table.get("ref") == "A1:D6"
    assert [column.get("name") for column in table.findall(f".//{{{MAIN_NS}}}tableColumn")] == [
        "Item",
        "Quantity",
        "UnitPrice",
        "LineTotal",
    ]


def test_f07_has_shared_groups_tamper_caches_table_and_merge(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F07"]
    cells = _cells(path, "FormulaBlocks")

    first_master = cells["C2"].find(f"{{{MAIN_NS}}}f")
    assert first_master is not None
    assert first_master.attrib == {"t": "shared", "ref": "C2:C11", "si": "0"}
    assert first_master.text == "A2*B2"
    for row in range(3, 12):
        follower = cells[f"C{row}"].find(f"{{{MAIN_NS}}}f")
        assert follower is not None
        assert follower.attrib == {"t": "shared", "si": "0"}
        assert follower.text is None

    tamper = cells["C12"].find(f"{{{MAIN_NS}}}f")
    assert tamper is not None
    assert tamper.get("t") is None
    assert tamper.text == "A12+B12"

    second_master = cells["C13"].find(f"{{{MAIN_NS}}}f")
    assert second_master is not None
    assert second_master.attrib == {"t": "shared", "ref": "C13:C21", "si": "1"}
    assert second_master.text == "A13*B13"
    for row in range(14, 22):
        follower = cells[f"C{row}"].find(f"{{{MAIN_NS}}}f")
        assert follower is not None
        assert follower.attrib == {"t": "shared", "si": "1"}
        assert follower.text is None

    assert {
        ref: float(cells[ref].find(f"{{{MAIN_NS}}}v").text)
        for ref in (f"C{row}" for row in range(2, 22))
    } == {
        "C2": 4.0,
        "C3": 10.0,
        "C4": 18.0,
        "C5": 8.0,
        "C6": 15.0,
        "C7": 24.0,
        "C8": 35.0,
        "C9": 48.0,
        "C10": 18.0,
        "C11": 30.0,
        "C12": 15.0,
        "C13": 60.0,
        "C14": 78.0,
        "C15": 28.0,
        "C16": 45.0,
        "C17": 64.0,
        "C18": 85.0,
        "C19": 108.0,
        "C20": 38.0,
        "C21": 60.0,
    }

    table = _xml_member(path, "xl/tables/table1.xml")
    assert table.get("name") == "FormulaBlocksTable"
    assert table.get("ref") == "A1:C21"
    worksheet = _worksheet(path, "FormulaBlocks")
    merge = worksheet.find(f".//{{{MAIN_NS}}}mergeCell")
    assert merge is not None
    assert merge.get("ref") == "E1:F1"


def test_f08_has_every_specified_error_cache_plus_unknown_typed_error(
    generated_paths: dict[str, Path],
) -> None:
    cells = _cells(generated_paths["F08"], "Errors")
    expected = (
        "#REF!",
        "#DIV/0!",
        "#N/A",
        "#VALUE!",
        "#NAME?",
        "#NUM!",
        "#SPILL!",
        "#CALC!",
        "#BLOCKED!",
        "#FIELD!",
    )
    assert tuple(_formula_and_value(cells[f"B{row}"]) for row in range(2, 12)) == tuple(
        ("NA()", value) for value in expected
    )
    assert {cells[f"B{row}"].get("t") for row in range(2, 12)} == {"e"}


def test_f10_has_numeric_external_link_map_and_missing_local_target(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F10"]
    assert _formula_and_value(_cells(path, "External")["A2"]) == ("[1]Data!A1", "0")
    with OOXMLParser(path) as parser:
        assert dict(parser.metadata.external_links) == {1: "missing/linked-budget.xlsx"}


def test_f11_has_dynamic_indirect_offset_formulas_and_caches(
    generated_paths: dict[str, Path],
) -> None:
    cells = _cells(generated_paths["F11"], "DynamicRefs")
    assert {ref: _formula_and_value(cells[ref]) for ref in ("B2", "C2")} == {
        "B2": ('INDIRECT("A2")', "10"),
        "C2": ("OFFSET(A2,1,0)", "20"),
    }


def test_f18_has_now_and_rand_volatile_formulas_and_caches(
    generated_paths: dict[str, Path],
) -> None:
    cells = _cells(generated_paths["F18"], "Volatile")
    assert {ref: _formula_and_value(cells[ref]) for ref in ("B2", "B3")} == {
        "B2": ("NOW()", "45292.5"),
        "B3": ("RAND()", "0.25"),
    }


def test_f02_has_three_islands_and_one_intentional_blank_row(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F02"]
    cells = _cells(path, "Islands")
    expected_refs = (
        _rect_refs(1, 1, 3, 3)
        | _rect_refs(1, 5, 3, 6)
        | _rect_refs(6, 2, 8, 6)
        | _rect_refs(2, 10, 4, 13)
    )

    assert set(cells) == expected_refs
    assert not any(ref.endswith("4") for ref in cells if ref[0] in "ABC")
    assert [_cell_text(cells[ref]) for ref in ("A1", "B1", "C1")] == [
        "Product",
        "Units",
        "UnitPrice",
    ]
    assert all(cells[ref].get("s") is not None for ref in ("A1", "F2", "B10"))


def test_f03_has_exact_cross_sheet_chains_tables_and_caches(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F03"]
    workbook = _xml_member(path, "xl/workbook.xml")
    assert [sheet.get("name") for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet")] == [
        "Inputs",
        "Calc",
        "Summary",
    ]
    assert _tables(path) == {
        ("InputsTable", "A1:C5"),
        ("CalcTable", "A1:D6"),
        ("SummaryTable", "A1:C10"),
    }

    calc = _cells(path, "Calc")
    assert {
        ref: _formula_and_value(cell)
        for ref, cell in calc.items()
        if _formula_and_value(cell)[0] is not None
    } == {
        "B2": ("Inputs!$B$3", "1000"),
        "C2": ("B2*Inputs!$B$4", "600"),
        "D2": ("(B2-C2)*(1-Inputs!$B$5)", "300"),
        "B3": ("B2*(1+Inputs!$B$2)", "1100"),
        "C3": ("B3*Inputs!$B$4", "660"),
        "D3": ("(B3-C3)*(1-Inputs!$B$5)", "330"),
        "B4": ("B3*(1+Inputs!$B$2)", "1210"),
        "C4": ("B4*Inputs!$B$4", "726"),
        "D4": ("(B4-C4)*(1-Inputs!$B$5)", "363"),
        "B5": ("B4*(1+Inputs!$B$2)", "1331"),
        "C5": ("B5*Inputs!$B$4", "798.6"),
        "D5": ("(B5-C5)*(1-Inputs!$B$5)", "399.3"),
        "B6": ("B5*(1+Inputs!$B$2)", "1464.1"),
        "C6": ("B6*Inputs!$B$4", "878.46"),
        "D6": ("(B6-C6)*(1-Inputs!$B$5)", "439.23"),
    }

    summary = _cells(path, "Summary")
    assert {
        ref: _formula_and_value(cell)
        for ref, cell in summary.items()
        if _formula_and_value(cell)[0] is not None
    } == {
        "C2": ("Calc!D2", "300"),
        "C3": ("Calc!D3", "330"),
        "C4": ("Calc!D4", "363"),
        "C5": ("Calc!D5", "399.3"),
        "C6": ("Calc!D6", "439.23"),
        "C7": ("Calc!B2", "1000"),
        "C8": ("Calc!B6", "1464.1"),
        "C9": ("SUM(Calc!C2:C6)", "3663.06"),
        "C10": ("SUM(Calc!D2:D6)", "1831.53"),
    }


def test_f04_has_global_and_sheet_scoped_names_used_by_formulas(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F04"]
    workbook = _xml_member(path, "xl/workbook.xml")
    defined_names = workbook.findall(f".//{{{MAIN_NS}}}definedName")
    assert [
        (name.get("name"), name.get("localSheetId"), "".join(name.itertext()))
        for name in defined_names
    ] == [
        ("BaseAmount", None, "'Inputs'!$B$2"),
        ("GlobalRate", None, "'Inputs'!$B$3"),
        ("ScopedRate", "1", "'Calc'!$B$2"),
    ]

    assert _formula_and_value(_cells(path, "Inputs")["B4"]) == (
        "BaseAmount*(1+GlobalRate)",
        "110",
    )
    assert {ref: _formula_and_value(_cells(path, "Calc")[ref]) for ref in ("B3", "B4")} == {
        "B3": ("BaseAmount", "100"),
        "B4": ("BaseAmount*(1+ScopedRate)", "105"),
    }

    with OOXMLParser(path) as parser:
        assert [
            (name.name, name.scope_sheet_order, name.refers_to, name.kind)
            for name in parser.metadata.defined_names
        ] == [
            ("BaseAmount", None, "'Inputs'!$B$2", "range"),
            ("GlobalRate", None, "'Inputs'!$B$3", "range"),
            ("ScopedRate", 1, "'Calc'!$B$2", "range"),
        ]


def test_f05_has_current_row_table_column_and_totals_row_formulas(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F05"]
    cells = _cells(path, "Structured")
    assert set(cells) == (_rect_refs(1, 1, 4, 6) - {"C6"} | {"F1", "F2"})
    assert {ref: _formula_and_value(cells[ref]) for ref in ("D2", "D3", "D4", "D5")} == {
        "D2": ("[@Qty]*[@Price]", "7.0"),
        "D3": ("[@Qty]*[@Price]", "6.25"),
        "D4": ("[@Qty]*[@Price]", "14.0"),
        "D5": ("[@Qty]*[@Price]", "27.0"),
    }
    assert {ref: _formula_and_value(cells[ref]) for ref in ("B6", "D6", "F2")} == {
        "B6": ("SUBTOTAL(109,Table1[Qty])", "18"),
        "D6": ("SUBTOTAL(109,Table1[LineTotal])", "54.25"),
        "F2": ("SUM(Table1[LineTotal])", "54.25"),
    }

    table = _xml_member(path, "xl/tables/table1.xml")
    assert table.get("name") == table.get("displayName") == "Table1"
    assert table.get("ref") == "A1:D6"
    assert table.get("headerRowCount") == "1"
    assert table.get("totalsRowCount") == "1"
    assert [
        (
            column.get("name"),
            column.get("totalsRowLabel"),
            column.get("totalsRowFunction"),
        )
        for column in table.findall(f".//{{{MAIN_NS}}}tableColumn")
    ] == [
        ("Product", "Totals", None),
        ("Qty", None, "sum"),
        ("Price", None, None),
        ("LineTotal", None, "sum"),
    ]


def test_f09a_has_one_small_two_cell_cycle_with_zero_caches(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F09a"]
    cells = _cells(path, "Circular")
    assert set(cells) == _rect_refs(1, 1, 2, 3)
    assert {ref: _formula_and_value(cells[ref]) for ref in ("B2", "B3")} == {
        "B2": ("B3+1", "0"),
        "B3": ("B2+1", "0"),
    }


def test_f09b_has_50000_running_total_formulas_with_zero_caches(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F09b"]
    cells = _cells(path, "RunningTotal")
    assert len(cells) == 100_004
    assert set(cells) == _rect_refs(1, 1, 2, 50_002)
    assert _formula_and_value(cells["B2"]) == (None, "0")
    assert {ref: _formula_and_value(cells[ref]) for ref in ("B3", "B25002", "B50002")} == {
        "B3": ("SUM($B$2:B2)", "0"),
        "B25002": ("SUM($B$2:B25001)", "0"),
        "B50002": ("SUM($B$2:B50001)", "0"),
    }
    assert all(
        _formula_and_value(cells[f"B{row}"]) == (f"SUM($B$2:B{row - 1})", "0")
        for row in range(3, 50_003)
    )


def test_f12_has_two_row_merged_header_hierarchy(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F12"]
    worksheet = _worksheet(path, "MergedHeaders")
    cells = _cells(path, "MergedHeaders")
    merges = {merge.get("ref") for merge in worksheet.findall(f".//{{{MAIN_NS}}}mergeCell")}

    assert merges == {"A1:A2", "B1:D1", "E1:F1"}
    assert [_cell_text(cells[ref]) for ref in ("A1", "B1", "E1")] == [
        "Region",
        "Revenue",
        "Units",
    ]
    assert [_cell_text(cells[ref]) for ref in ("B2", "C2", "D2", "E2", "F2")] == [
        "Q1",
        "Q2",
        "Q3",
        "Actual",
        "Target",
    ]
    assert len(cells) == 32


def test_f13_stores_raw_date_serials_and_all_mixed_type_styles(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F13"]
    cells = _cells(path, "MixedTypes")
    styles = _xml_member(path, "xl/styles.xml")
    cell_xfs = styles.find(f"{{{MAIN_NS}}}cellXfs")
    assert cell_xfs is not None
    xfs = cell_xfs.findall(f"{{{MAIN_NS}}}xf")
    num_fmts = {
        item.get("numFmtId"): item.get("formatCode")
        for item in styles.findall(f".//{{{MAIN_NS}}}numFmt")
    }

    assert [_formula_and_value(cells[f"B{row}"])[1] for row in range(2, 8)] == [
        "45292",
        "45323",
        "45352",
        "45383",
        "45413",
        "45444",
    ]
    assert xfs[int(cells["B2"].get("s"))].get("numFmtId") == "164"
    assert xfs[int(cells["B5"].get("s"))].get("numFmtId") == "14"
    assert num_fmts == {
        "164": "yyyy-mm-dd",
        "165": '"$"#,##0.00',
        "166": "0.0%",
    }
    assert cells["E2"].get("t") == "inlineStr"
    assert _cell_text(cells["E2"]) == "00100"
    assert cells["F2"].get("t") == "b"
    assert cells["G2"].get("t") == "n"
    assert cells["G3"].get("t") == "inlineStr"
    assert _tables(path) == {("MixedTypesTable", "A1:G7")}


def test_f14_has_empty_sheets_and_two_sparse_singletons(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F14"]
    workbook = _xml_member(path, "xl/workbook.xml")
    assert [sheet.get("name") for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet")] == [
        "EmptyBefore",
        "LoneCells",
        "EmptyAfter",
    ]
    assert _cells(path, "EmptyBefore") == {}
    assert set(_cells(path, "LoneCells")) == {"B2", "X100"}
    assert _cells(path, "EmptyAfter") == {}
    dimension = _worksheet(path, "LoneCells").find(f"{{{MAIN_NS}}}dimension")
    assert dimension is not None
    assert dimension.get("ref") == "B2:X100"


def test_f15_has_exact_three_dimensional_formula_span_and_cache(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F15"]
    workbook = _xml_member(path, "xl/workbook.xml")
    assert [sheet.get("name") for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet")] == [
        "Jan",
        "Feb",
        "Mar",
        "Summary",
    ]
    assert {
        sheet: _formula_and_value(_cells(path, sheet)["B2"])[1] for sheet in ("Jan", "Feb", "Mar")
    } == {"Jan": "10", "Feb": "20", "Mar": "30"}
    assert _formula_and_value(_cells(path, "Summary")["B2"]) == (
        "SUM(Jan:Mar!B2)",
        "60",
    )


def test_f20_has_40_sheets_12_islands_and_300_typed_names(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F20"]
    workbook = _xml_member(path, "xl/workbook.xml")
    sheets = workbook.findall(f".//{{{MAIN_NS}}}sheet")
    defined_names = workbook.findall(f".//{{{MAIN_NS}}}definedName")

    assert [sheet.get("name") for sheet in sheets] == [
        f"Stress{number:02d}" for number in range(1, 41)
    ]
    assert sheets[-2].get("state") == "hidden"
    assert sheets[-1].get("state") == "veryHidden"
    assert all(sheet.get("state") == "visible" for sheet in sheets[:-2])

    stress_cells = _cells(path, "Stress01")
    expected_refs: set[str] = set()
    for block_number in range(12):
        first_column = 1 + block_number * 4
        expected_refs |= _rect_refs(first_column, 1, first_column + 1, 14 - block_number)
    assert set(stress_cells) == expected_refs
    assert len(stress_cells) == 204
    assert all(_cells(path, f"Stress{number:02d}") == {} for number in range(2, 41))

    assert len(defined_names) == 300
    assert Counter(name.get("name", "")[-1] for name in defined_names) == {
        "R": 60,
        "M": 60,
        "C": 60,
        "F": 60,
        "L": 60,
    }
    assert [(name.get("name"), "".join(name.itertext())) for name in defined_names[:5]] == [
        ("N001R", "'Stress01'!$A$2:$B$3"),
        ("N001M", "'Stress01'!$A$2:$A$3,'Stress01'!$E$2:$E$3"),
        ("N001C", "=1"),
        ("N001F", "=SUM('Stress01'!$B$2:$B$3)+1"),
        ("N001L", "=_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)"),
    ]

    with OOXMLParser(path) as parser:
        assert Counter(name.kind for name in parser.metadata.defined_names) == {
            "range": 60,
            "multi_range": 60,
            "constant": 60,
            "formula": 60,
            "lambda": 60,
        }


def test_f19_has_exact_modern_formulas_names_spills_and_caches(
    generated_paths: dict[str, Path],
) -> None:
    path = generated_paths["F19"]
    cells = _cells(path, "Modern")

    assert {
        ref: _formula_and_value(cells[ref]) for ref in ("A1", "B1", "C1", "D1", "E1", "F1", "G2")
    } == {
        "A1": ("_xlfn._xlws.FILTER(I2:I4,I2:I4>=20)", "20"),
        "B1": ("SUM(A1#)", "50"),
        "C1": ("SUM(FilteredValues#)", "50"),
        "D1": (
            "_xlfn.LET(_xlpm.rate,I2,_xlpm.bonus,1,_xlpm.rate*3+_xlpm.bonus)",
            "31",
        ),
        "E1": ("DoubleIt(I3)", "40"),
        "F1": ('_xlfn.XLOOKUP("beta",H2:H4,I2:I4,"missing")', "20"),
        "G2": ("@I2:I4", "10"),
    }
    assert _formula_and_value(cells["A2"]) == (None, "30")

    workbook = _xml_member(path, "xl/workbook.xml")
    defined_names = workbook.findall(f".//{{{MAIN_NS}}}definedName")
    assert [(name.get("name"), "".join(name.itertext())) for name in defined_names] == [
        ("DoubleIt", "=_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)"),
        ("FilteredValues", "'Modern'!$A$1"),
    ]

    with OOXMLParser(path) as parser:
        names = parser.metadata.defined_names
        assert [(name.name, name.refers_to, name.kind) for name in names] == [
            ("DoubleIt", "=_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)", "lambda"),
            ("FilteredValues", "'Modern'!$A$1", "range"),
        ]
        assert names[0].areas == ()
        assert [
            (
                area.sheet_name,
                area.rect.row_min,
                area.rect.row_max,
                area.rect.col_min,
                area.rect.col_max,
            )
            for area in names[1].areas
        ] == [("Modern", 1, 1, 1, 1)]
