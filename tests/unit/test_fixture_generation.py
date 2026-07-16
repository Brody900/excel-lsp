from __future__ import annotations

import runpy
from collections.abc import Callable
from pathlib import Path
from typing import Any, cast
from zipfile import ZipFile

import lxml.etree as etree

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
FIXED_ZIP_TIMESTAMP = (1980, 1, 1, 0, 0, 0)
GenerateAll = Callable[[Path], dict[str, Path]]
generate_all = cast(
    GenerateAll,
    runpy.run_path(str(Path(__file__).parents[1] / "fixtures" / "generate.py"))["generate_all"],
)


def _xml_member(path: Path, member: str):
    with ZipFile(path) as archive:
        return etree.fromstring(archive.read(member))


def _cells(path: Path) -> dict[str, Any]:
    worksheet = _xml_member(path, "xl/worksheets/sheet1.xml")
    return {
        cell.get("r"): cell for cell in worksheet.findall(f".//{{{MAIN_NS}}}c") if cell.get("r")
    }


def test_generation_is_byte_identical_with_stable_zip_metadata(tmp_path: Path) -> None:
    first_paths = generate_all(tmp_path)
    first_bytes = {fixture_id: path.read_bytes() for fixture_id, path in first_paths.items()}

    second_paths = generate_all(tmp_path)

    assert second_paths.keys() == first_paths.keys() == {"F01", "F07"}
    for fixture_id, path in second_paths.items():
        assert path.read_bytes() == first_bytes[fixture_id]
        with ZipFile(path) as archive:
            infos = archive.infolist()
        assert [info.filename for info in infos] == sorted(info.filename for info in infos)
        assert {info.date_time for info in infos} == {FIXED_ZIP_TIMESTAMP}


def test_f01_has_listobject_formulas_and_injected_caches(tmp_path: Path) -> None:
    path = generate_all(tmp_path)["F01"]
    cells = _cells(path)
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


def test_f07_has_shared_groups_tamper_caches_table_and_merge(tmp_path: Path) -> None:
    path = generate_all(tmp_path)["F07"]
    cells = _cells(path)

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
    worksheet = _xml_member(path, "xl/worksheets/sheet1.xml")
    merge = worksheet.find(f".//{{{MAIN_NS}}}mergeCell")
    assert merge is not None
    assert merge.get("ref") == "E1:F1"
