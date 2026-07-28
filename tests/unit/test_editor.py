from __future__ import annotations

import json
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, cast
from zipfile import ZIP_DEFLATED, ZipFile

import lxml.etree as etree
import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table

import excel_lsp.core.edit.service as edit_service
import excel_lsp.core.edit.writer as edit_writer
from excel_lsp.core.edit import (
    CellEdit,
    WriteScalar,
    patch_workbook,
    set_column_formula,
    write_cells,
)
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.index.lifecycle import index_workbook
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import Rect
from excel_lsp.core.parse import OOXMLParser
from excel_lsp.core.parse._xml import attr_by_local, child_by_local, local_name, parse_xml

FIXTURES = Path(__file__).parents[1] / "fixtures" / "generated"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_TYPE_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _copy_fixture(tmp_path: Path, name: str) -> Path:
    destination = tmp_path / name
    shutil.copyfile(FIXTURES / name, destination)
    return destination


def _workbook_hash(path: Path) -> str:
    with OOXMLParser(path) as parser:
        return parser.hashes.whole_file


def _parts(path: Path) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {
            info.filename: archive.read(info) for info in archive.infolist() if not info.is_dir()
        }


def _worksheet_root(path: Path, part: str = "xl/worksheets/sheet1.xml") -> etree._Element:
    with ZipFile(path) as archive:
        return parse_xml(archive.read(part))


def _cells(root: etree._Element) -> dict[str, etree._Element]:
    return {
        str(cell.get("r")): cell
        for cell in root.iter()
        if isinstance(cell.tag, str) and local_name(cell.tag) == "c" and cell.get("r")
    }


def _rewrite_parts(path: Path, replacements: dict[str, bytes]) -> None:
    temporary = path.with_suffix(".rewrite")
    with (
        ZipFile(path, "r") as source,
        ZipFile(
            temporary,
            "w",
            compression=ZIP_DEFLATED,
        ) as target,
    ):
        seen: set[str] = set()
        for info in source.infolist():
            if info.filename in replacements:
                target.writestr(info, replacements[info.filename])
                seen.add(info.filename)
            else:
                target.writestr(info, source.read(info) if not info.is_dir() else b"")
        for name, payload in replacements.items():
            if name not in seen:
                target.writestr(name, payload)
    temporary.replace(path)


def _inject_calc_chain(path: Path) -> None:
    parts = _parts(path)
    relationships = parse_xml(parts["xl/_rels/workbook.xml.rels"])
    etree.SubElement(
        relationships,
        f"{{{PACKAGE_REL_NS}}}Relationship",
        Id="rIdCalcChain",
        Type=f"{REL_TYPE_BASE}/calcChain",
        Target="calcChain.xml",
    )
    content_types = parse_xml(parts["[Content_Types].xml"])
    etree.SubElement(
        content_types,
        f"{{{CONTENT_TYPES_NS}}}Override",
        PartName="/xl/calcChain.xml",
        ContentType=("application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"),
    )
    calc_chain = etree.Element(f"{{{MAIN_NS}}}calcChain")
    etree.SubElement(calc_chain, f"{{{MAIN_NS}}}c", r="D2", i="1")
    _rewrite_parts(
        path,
        {
            "xl/_rels/workbook.xml.rels": etree.tostring(relationships),
            "[Content_Types].xml": etree.tostring(content_types),
            "xl/calcChain.xml": etree.tostring(calc_chain),
        },
    )


def _inject_array_formula(path: Path) -> None:
    root = _worksheet_root(path)
    cells = _cells(root)
    anchor = cells["D2"]
    formula = child_by_local(anchor, "f")
    assert formula is not None
    formula.set("t", "array")
    formula.set("ref", "D2:D3")
    _rewrite_parts(path, {"xl/worksheets/sheet1.xml": etree.tostring(root)})


def test_surgical_write_changes_only_declared_parts_and_serializes_cell_types(
    tmp_path: Path,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    before = _parts(workbook)

    result = patch_workbook(
        workbook,
        (
            CellEdit.value("Sales", "A2", "  spaced  "),
            CellEdit.value("Sales", "B2", 7),
            CellEdit.value("Sales", "C2", True),
            CellEdit.formula("Sales", "D2", "=B2*2"),
            CellEdit.value("Sales", "F10", None),
        ),
        expected_workbook_hash=_workbook_hash(workbook),
    )

    after = _parts(workbook)
    assert result.modified_parts == ("xl/worksheets/sheet1.xml",)
    assert result.deleted_parts == ()
    assert result.workbook_hash_before != result.workbook_hash_after
    assert {cell.ref for cell in result.patched_cells} == {"A2", "B2", "C2", "D2", "F10"}
    for part, payload in before.items():
        if part not in result.modified_parts:
            assert after[part] == payload

    root = _worksheet_root(workbook)
    cells = _cells(root)
    assert cells["A2"].get("t") == "inlineStr"
    text = next(element for element in cells["A2"].iter() if local_name(element.tag) == "t")
    assert text.text == "  spaced  "
    assert text.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    assert cells["B2"].get("t") is None
    assert child_by_local(cells["B2"], "v").text == "7"  # type: ignore[union-attr]
    assert cells["C2"].get("t") == "b"
    assert child_by_local(cells["C2"], "v").text == "1"  # type: ignore[union-attr]
    assert child_by_local(cells["D2"], "f").text == "B2*2"  # type: ignore[union-attr]
    assert child_by_local(cells["D2"], "v") is None
    assert child_by_local(cells["F10"], "f") is None
    assert child_by_local(cells["F10"], "v") is None
    assert child_by_local(cells["F10"], "is") is None
    dimension = child_by_local(root, "dimension")
    assert dimension is not None
    assert dimension.get("ref") == "A1:F10"

    workbook_root = parse_xml(after["xl/workbook.xml"])
    calc = child_by_local(workbook_root, "calcPr")
    assert calc is not None
    assert attr_by_local(calc, "fullCalcOnLoad") == "1"


def test_cell_payloads_remain_before_preserved_extension_lists(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    root = _worksheet_root(workbook)
    cells = _cells(root)
    for ref in ("A2", "B2", "D2"):
        namespace = cells[ref].tag[1:].split("}", 1)[0]
        extension_list = etree.SubElement(cells[ref], f"{{{namespace}}}extLst")
        etree.SubElement(
            extension_list,
            f"{{{namespace}}}ext",
            uri="{25C4219D-8188-4D27-B50F-8B7A0A88E7D6}",
        )
    _rewrite_parts(workbook, {"xl/worksheets/sheet1.xml": etree.tostring(root)})

    patch_workbook(
        workbook,
        (
            CellEdit.value("Sales", "A2", "replacement"),
            CellEdit.value("Sales", "B2", 7),
            CellEdit.formula("Sales", "D2", "=B2*C2"),
        ),
        expected_workbook_hash=_workbook_hash(workbook),
    )

    updated = _cells(_worksheet_root(workbook))
    assert [local_name(child.tag) for child in updated["A2"]] == ["is", "extLst"]
    assert [local_name(child.tag) for child in updated["B2"]] == ["v", "extLst"]
    assert [local_name(child.tag) for child in updated["D2"]] == ["f", "extLst"]
    with OOXMLParser(workbook) as parser:
        parser.parse_sheet(parser.metadata.sheets[0], lambda _cell: None)


@pytest.mark.parametrize(
    "successor",
    (
        "oleSize",
        "customWorkbookViews",
        "pivotCaches",
        "smartTagPr",
        "smartTagTypes",
        "webPublishing",
        "fileRecoveryPr",
        "webPublishObjects",
        "extLst",
    ),
)
def test_new_calc_properties_precede_every_later_workbook_child(
    tmp_path: Path,
    successor: str,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    parts = _parts(workbook)
    root = parse_xml(parts["xl/workbook.xml"])
    calc = child_by_local(root, "calcPr")
    assert calc is not None
    root.remove(calc)
    namespace = root.tag[1:].split("}", 1)[0]
    etree.SubElement(root, f"{{{namespace}}}{successor}")
    names_before = [local_name(child.tag) for child in root]
    successor_index = names_before.index(successor)
    expected_names = [
        *names_before[:successor_index],
        "calcPr",
        *names_before[successor_index:],
    ]
    _rewrite_parts(workbook, {"xl/workbook.xml": etree.tostring(root)})

    patch_workbook(
        workbook,
        (CellEdit.value("Sales", "B2", 7),),
        expected_workbook_hash=_workbook_hash(workbook),
    )

    updated = parse_xml(_parts(workbook)["xl/workbook.xml"])
    assert [local_name(child.tag) for child in updated] == expected_names
    updated_calc = child_by_local(updated, "calcPr")
    assert updated_calc is not None
    assert attr_by_local(updated_calc, "fullCalcOnLoad") == "1"
    with OOXMLParser(workbook) as parser:
        assert parser.metadata.sheets[0].name == "Sales"


def test_within_extent_write_drops_advisory_dimension_and_preserves_order_and_style(
    tmp_path: Path,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    original_root = _worksheet_root(workbook)
    original_style = _cells(original_root)["A2"].get("s")

    patch_workbook(
        workbook,
        (CellEdit.value("Sales", "A2", "replacement"),),
        expected_workbook_hash=_workbook_hash(workbook),
    )

    root = _worksheet_root(workbook)
    assert child_by_local(root, "dimension") is None
    assert _cells(root)["A2"].get("s") == original_style
    sheet_data = child_by_local(root, "sheetData")
    assert sheet_data is not None
    row_numbers = [int(row.get("r")) for row in sheet_data if local_name(row.tag) == "row"]
    assert row_numbers == sorted(row_numbers)
    for row in sheet_data:
        refs = [cell.get("r") for cell in row if local_name(cell.tag) == "c"]
        assert refs == sorted(refs, key=lambda ref: _ref_key(str(ref)))


def _ref_key(ref: str) -> tuple[int, int]:
    from excel_lsp.core.parse.coordinates import parse_cell_ref

    return parse_cell_ref(ref)


def test_editing_shared_follower_expands_the_complete_group_before_write(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "formula_blocks.xlsx")
    result = patch_workbook(
        workbook,
        (CellEdit.value("FormulaBlocks", "C5", 999),),
        expected_workbook_hash=_workbook_hash(workbook),
    )

    cells = _cells(_worksheet_root(workbook))
    expected_formulas = {row: f"A{row}*B{row}" for row in range(2, 12) if row != 5}
    for row, expected in expected_formulas.items():
        formula = child_by_local(cells[f"C{row}"], "f")
        assert formula is not None
        assert formula.text == expected
        assert formula.get("t") is None
        assert formula.get("si") is None
        assert formula.get("ref") is None
    assert child_by_local(cells["C5"], "f") is None
    assert child_by_local(cells["C5"], "v").text == "999"  # type: ignore[union-attr]
    untouched_group = child_by_local(cells["C13"], "f")
    assert untouched_group is not None
    assert untouched_group.get("t") == "shared"
    assert untouched_group.get("si") == "1"
    assert {cell.ref for cell in result.patched_cells} == {
        *(f"C{row}" for row in range(2, 12)),
    }
    assert sum(cell.requested for cell in result.patched_cells) == 1


def test_multi_cell_array_formula_refuses_anchor_and_follower_without_writing(
    tmp_path: Path,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    _inject_array_formula(workbook)
    before = workbook.read_bytes()
    expected_hash = _workbook_hash(workbook)

    for ref in ("D2", "D3"):
        with pytest.raises(ExcelLSPError) as captured:
            patch_workbook(
                workbook,
                (CellEdit.value("Sales", ref, 1),),
                expected_workbook_hash=expected_hash,
            )
        assert captured.value.code is ErrorCode.ARRAY_FORMULA
        assert workbook.read_bytes() == before


def test_calc_chain_part_relationship_and_override_are_deleted_together(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    _inject_calc_chain(workbook)
    before = _parts(workbook)

    result = patch_workbook(
        workbook,
        (CellEdit.value("Sales", "B2", 8),),
        expected_workbook_hash=_workbook_hash(workbook),
    )

    after = _parts(workbook)
    assert result.deleted_parts == ("xl/calcChain.xml",)
    assert "xl/calcChain.xml" not in after
    assert {
        "[Content_Types].xml",
        "xl/_rels/workbook.xml.rels",
        "xl/worksheets/sheet1.xml",
    } == set(result.modified_parts)
    for part, payload in before.items():
        if part not in result.modified_parts and part not in result.deleted_parts:
            assert after[part] == payload
    relationships = parse_xml(after["xl/_rels/workbook.xml.rels"])
    assert all(
        not (attr_by_local(relationship, "Type") or "").endswith("/calcChain")
        for relationship in relationships
    )
    content_types = parse_xml(after["[Content_Types].xml"])
    assert all(
        attr_by_local(override, "PartName") != "/xl/calcChain.xml" for override in content_types
    )


def test_lockfile_precedes_conflict_and_neither_failure_mutates_workbook(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    before = workbook.read_bytes()
    lockfile = workbook.with_name(f"~${workbook.name}")
    lockfile.write_bytes(b"locked")

    with pytest.raises(ExcelLSPError) as locked:
        patch_workbook(
            workbook,
            (CellEdit.value("Sales", "A2", "x"),),
            expected_workbook_hash="0" * 64,
        )
    assert locked.value.code is ErrorCode.OPEN_IN_EXCEL
    assert workbook.read_bytes() == before

    lockfile.unlink()
    with pytest.raises(ExcelLSPError) as conflict:
        patch_workbook(
            workbook,
            (CellEdit.value("Sales", "A2", "x"),),
            expected_workbook_hash="0" * 64,
        )
    assert conflict.value.code is ErrorCode.CONFLICT
    assert workbook.read_bytes() == before


def test_replace_retry_preserves_intervening_external_workbook(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    expected_hash = _workbook_hash(workbook)
    external_bytes = b"intervening external workbook bytes"
    real_replace = edit_writer.os.replace
    replace_calls = 0

    def fail_first_replace(source: Any, destination: Any) -> None:
        nonlocal replace_calls
        replace_calls += 1
        if replace_calls == 1:
            raise PermissionError("injected first replacement failure")
        real_replace(source, destination)

    def install_external_workbook(_delay: float) -> None:
        workbook.write_bytes(external_bytes)

    monkeypatch.setattr(edit_writer.os, "replace", fail_first_replace)
    monkeypatch.setattr(edit_writer.time, "sleep", install_external_workbook)

    with pytest.raises(ExcelLSPError) as captured:
        patch_workbook(
            workbook,
            (CellEdit.value("Sales", "B2", 7),),
            expected_workbook_hash=expected_hash,
        )

    assert captured.value.code is ErrorCode.CONFLICT
    assert replace_calls == 1
    assert workbook.read_bytes() == external_bytes


def test_replace_retry_rechecks_excel_lock_before_destination_hash(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    expected_hash = _workbook_hash(workbook)
    external_bytes = b"intervening locked workbook bytes"
    lockfile = workbook.with_name(f"~${workbook.name}")
    real_hash_file = edit_writer._hash_file
    hash_calls = 0
    replace_calls = 0

    def count_hash(path: Path) -> str:
        nonlocal hash_calls
        hash_calls += 1
        return real_hash_file(path)

    def refuse_replace(_source: Any, _destination: Any) -> None:
        nonlocal replace_calls
        replace_calls += 1
        raise PermissionError("injected first replacement failure")

    def open_in_excel(_delay: float) -> None:
        workbook.write_bytes(external_bytes)
        lockfile.write_bytes(b"locked")

    monkeypatch.setattr(edit_writer, "_hash_file", count_hash)
    monkeypatch.setattr(edit_writer.os, "replace", refuse_replace)
    monkeypatch.setattr(edit_writer.time, "sleep", open_in_excel)

    with pytest.raises(ExcelLSPError) as captured:
        patch_workbook(
            workbook,
            (CellEdit.value("Sales", "B2", 7),),
            expected_workbook_hash=expected_hash,
        )

    assert captured.value.code is ErrorCode.OPEN_IN_EXCEL
    assert replace_calls == 1
    assert hash_calls == 2
    assert workbook.read_bytes() == external_bytes


def test_replace_retry_succeeds_when_preconditions_remain_unchanged(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    expected_hash = _workbook_hash(workbook)
    real_replace = edit_writer.os.replace
    replace_calls = 0

    def fail_once_then_replace(source: Any, destination: Any) -> None:
        nonlocal replace_calls
        replace_calls += 1
        if replace_calls == 1:
            raise PermissionError("injected first replacement failure")
        real_replace(source, destination)

    monkeypatch.setattr(edit_writer.os, "replace", fail_once_then_replace)
    monkeypatch.setattr(edit_writer.time, "sleep", lambda _delay: None)

    result = patch_workbook(
        workbook,
        (CellEdit.value("Sales", "B2", 7),),
        expected_workbook_hash=expected_hash,
    )

    assert replace_calls == 2
    assert result.workbook_hash_after != expected_hash
    value = child_by_local(_cells(_worksheet_root(workbook))["B2"], "v")
    assert value is not None
    assert value.text == "7"


def test_repeated_replace_failure_returns_locked_without_mutation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    before = workbook.read_bytes()
    replace_calls = 0

    def refuse_replace(_source: Any, _destination: Any) -> None:
        nonlocal replace_calls
        replace_calls += 1
        raise PermissionError("injected replacement failure")

    monkeypatch.setattr(edit_writer.os, "replace", refuse_replace)
    monkeypatch.setattr(edit_writer.time, "sleep", lambda _delay: None)

    with pytest.raises(ExcelLSPError) as captured:
        patch_workbook(
            workbook,
            (CellEdit.value("Sales", "B2", 7),),
            expected_workbook_hash=_workbook_hash(workbook),
        )

    assert captured.value.code is ErrorCode.LOCKED
    assert replace_calls == 2
    assert workbook.read_bytes() == before


@pytest.mark.parametrize(
    "edit",
    (
        CellEdit.value("Sales", "A2", cast(WriteScalar, datetime(2026, 7, 28))),
        CellEdit.value("Sales", "A2", float("nan")),
        CellEdit.formula("Sales", "A2", "SUM(B2:B3)"),
    ),
)
def test_invalid_write_values_fail_before_any_package_mutation(
    tmp_path: Path,
    edit: CellEdit,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    before = workbook.read_bytes()
    with pytest.raises(ExcelLSPError) as captured:
        patch_workbook(
            workbook,
            (edit,),
            expected_workbook_hash=_workbook_hash(workbook),
        )
    assert captured.value.code is ErrorCode.INVALID_VALUE
    assert workbook.read_bytes() == before


@pytest.mark.parametrize(
    "edit",
    (
        CellEdit.value("Sales", "A2", "x" * 32_768),
        CellEdit.value("Sales", "A2", "😀" * 16_384),
        CellEdit.value("Sales", "A2", "\ud800"),
        CellEdit.value("Sales", "A2", cast(WriteScalar, 10**400)),
        CellEdit.formula("Sales", "A2", "=" + "A" * 8_192),
        CellEdit.formula("Sales", "A2", "==1"),
        CellEdit.formula("Sales", "A2", "=1\x01+1"),
    ),
)
def test_excel_text_and_numeric_boundaries_reject_without_mutation(
    tmp_path: Path,
    edit: CellEdit,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    before = workbook.read_bytes()

    with pytest.raises(ExcelLSPError) as captured:
        patch_workbook(
            workbook,
            (edit,),
            expected_workbook_hash=_workbook_hash(workbook),
        )

    assert captured.value.code is ErrorCode.INVALID_VALUE
    assert workbook.read_bytes() == before


def test_integral_float_serializes_without_a_decimal_suffix(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")

    patch_workbook(
        workbook,
        (CellEdit.value("Sales", "B2", 7.0),),
        expected_workbook_hash=_workbook_hash(workbook),
    )

    value = child_by_local(_cells(_worksheet_root(workbook))["B2"], "v")
    assert value is not None
    assert value.text == "7"


def test_duplicate_and_oversized_edit_sets_are_rejected(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    expected_hash = _workbook_hash(workbook)
    with pytest.raises(ExcelLSPError) as duplicate:
        patch_workbook(
            workbook,
            (
                CellEdit.value("Sales", "A2", 1),
                CellEdit.formula("sales", "$A$2", "=1"),
            ),
            expected_workbook_hash=expected_hash,
        )
    assert duplicate.value.code is ErrorCode.INVALID_REF

    with pytest.raises(ExcelLSPError) as oversized:
        patch_workbook(
            workbook,
            tuple(CellEdit.value("Sales", f"A{row}", row) for row in range(1, 502)),
            expected_workbook_hash=expected_hash,
        )
    assert oversized.value.code is ErrorCode.INVALID_VALUE


def test_write_service_directly_patches_index_and_marks_transitive_dependents_stale(
    tmp_path: Path,
) -> None:
    workbook = _copy_fixture(tmp_path, "cross_sheet_model.xlsx")
    opened = index_workbook(workbook)

    result = write_cells(workbook, (CellEdit.value("Inputs", "B2", 0.2),))

    assert result.direct_index_patch is True
    assert result.generation == opened.generation + 1
    assert result.stale_blocks >= 2
    with IndexStore(opened.index_path) as store:
        row = store.connection.execute(
            "SELECT value, value_type, formula FROM cells "
            "WHERE sheet_id = 1 AND row = 2 AND col = 2"
        ).fetchone()
        assert row is not None
        assert row["value"] == 0.2
        assert row["value_type"] == "number"
        assert row["formula"] is None
        assert store.is_stale("Calc", Rect(3, 3, 2, 2))
        assert store.is_stale("Summary", Rect(3, 3, 3, 3))
        assert not store.is_stale("Inputs", Rect(2, 2, 2, 2))
        stale_report = store.get_diagnostics(code="I_STALE")
        assert stale_report.total == result.stale_blocks
        assert {item.sheet for item in stale_report.diagnostics} >= {"Calc", "Summary"}
        assert all(item.code == "I_STALE" for item in stale_report.diagnostics)
        with OOXMLParser(workbook) as parser:
            assert store.get_meta("workbook_hash") == parser.hashes.whole_file
        assert store.get_meta("mtime_ns") == str(workbook.stat().st_mtime_ns)
        assert store.get_meta("size") == str(workbook.stat().st_size)

    fresh = index_workbook(workbook)
    assert fresh.changed is False
    assert fresh.generation == result.generation


def test_formula_write_removes_cache_and_marks_the_written_formula_stale(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "cross_sheet_model.xlsx")
    result = write_cells(
        workbook,
        (CellEdit.formula("Summary", "C2", "=Calc!D2*2"),),
    )

    root = _worksheet_root(workbook, "xl/worksheets/sheet3.xml")
    cell = _cells(root)["C2"]
    assert child_by_local(cell, "f").text == "Calc!D2*2"  # type: ignore[union-attr]
    assert child_by_local(cell, "v") is None
    with IndexStore(result.patch.path.with_name(f"{result.patch.path.name}.xlsp.db")) as store:
        row = store.connection.execute(
            "SELECT value, value_type, formula FROM cells "
            "WHERE sheet_id = 3 AND row = 2 AND col = 3"
        ).fetchone()
        assert row is not None
        assert row["value"] is None
        assert row["value_type"] == "blank"
        assert row["formula"] == "=Calc!D2*2"
        assert store.is_stale("Summary", Rect(2, 2, 3, 3))


def test_null_write_deletes_semantic_cell_row_but_keeps_valid_empty_ooxml_cell(
    tmp_path: Path,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    result = write_cells(workbook, (CellEdit.value("Sales", "A2", None),))

    assert "A2" in _cells(_worksheet_root(workbook))
    with IndexStore(workbook.with_name(f"{workbook.name}.xlsp.db")) as store:
        row = store.connection.execute(
            "SELECT 1 FROM cells WHERE sheet_id = 1 AND row = 2 AND col = 1"
        ).fetchone()
        assert row is None
        assert store.generation == result.generation


def test_service_shared_group_expansion_updates_every_indexed_formula(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "formula_blocks.xlsx")
    result = write_cells(workbook, (CellEdit.value("FormulaBlocks", "C5", 999),))

    with IndexStore(workbook.with_name(f"{workbook.name}.xlsp.db")) as store:
        rows = store.connection.execute(
            "SELECT row, formula, formula_kind, shared_index FROM cells "
            "WHERE sheet_id = 1 AND col = 3 AND row BETWEEN 2 AND 11 ORDER BY row"
        ).fetchall()
        assert len(rows) == 10
        for row in rows:
            if row["row"] == 5:
                assert row["formula"] is None
                assert row["formula_kind"] is None
            else:
                assert row["formula"] == f"=A{row['row']}*B{row['row']}"
                assert row["formula_kind"] == "normal"
            assert row["shared_index"] is None
        assert store.generation == result.generation


def test_post_replace_direct_patch_failure_recovers_index_and_staleness(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = _copy_fixture(tmp_path, "cross_sheet_model.xlsx")

    def fail_direct_patch(*_args: object, **_kwargs: object) -> int:
        raise RuntimeError("injected direct patch failure")

    monkeypatch.setattr(IndexStore, "apply_editor_patch", fail_direct_patch)
    result = write_cells(workbook, (CellEdit.value("Inputs", "B2", 0.3),))

    assert result.direct_index_patch is False
    with IndexStore(workbook.with_name(f"{workbook.name}.xlsp.db")) as store:
        row = store.connection.execute(
            "SELECT value FROM cells WHERE sheet_id = 1 AND row = 2 AND col = 2"
        ).fetchone()
        assert row is not None
        assert row["value"] == 0.3
        assert store.is_stale("Calc", Rect(3, 3, 2, 2))


def test_post_replace_workbook_race_reconciles_index_and_returns_conflict(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = _copy_fixture(tmp_path, "cross_sheet_model.xlsx")
    real_patch = edit_service.patch_workbook

    def patch_then_race(*args: Any, **kwargs: Any) -> Any:
        result = real_patch(*args, **kwargs)
        root = _worksheet_root(workbook, "xl/worksheets/sheet2.xml")
        cell = _cells(root)["A2"]
        value = child_by_local(cell, "v")
        assert value is not None
        value.text = "9999"
        _rewrite_parts(workbook, {"xl/worksheets/sheet2.xml": etree.tostring(root)})
        return result

    monkeypatch.setattr(edit_service, "patch_workbook", patch_then_race)

    with pytest.raises(ExcelLSPError) as captured:
        write_cells(workbook, (CellEdit.value("Inputs", "B2", 0.3),))

    assert captured.value.code is ErrorCode.CONFLICT
    refreshed = index_workbook(workbook)
    assert refreshed.changed is False
    with IndexStore(refreshed.index_path) as store:
        indexed = store.connection.execute(
            "SELECT value FROM cells WHERE sheet_id = 2 AND row = 2 AND col = 1"
        ).fetchone()
        assert indexed is not None
        assert indexed["value"] == 9999


def test_race_during_sheet_collection_recovers_current_workbook_generation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = _copy_fixture(tmp_path, "cross_sheet_model.xlsx")
    external = tmp_path / "external-writer.xlsx"
    shutil.copyfile(FIXTURES / "cross_sheet_model.xlsx", external)
    external_root = _worksheet_root(external, "xl/worksheets/sheet2.xml")
    external_value = child_by_local(_cells(external_root)["A2"], "v")
    assert external_value is not None
    external_value.text = "9999"
    _rewrite_parts(
        external,
        {"xl/worksheets/sheet2.xml": etree.tostring(external_root)},
    )
    external_bytes = external.read_bytes()
    real_collect = edit_service._collect_sheet_patches
    collection_calls = 0

    def replace_during_collection(parser: Any, patch_result: Any) -> Any:
        nonlocal collection_calls
        collection_calls += 1
        workbook.write_bytes(external_bytes)
        return real_collect(parser, patch_result)

    monkeypatch.setattr(edit_service, "_collect_sheet_patches", replace_during_collection)

    with pytest.raises(ExcelLSPError) as captured:
        write_cells(workbook, (CellEdit.value("Inputs", "B2", 0.3),))

    assert captured.value.code is ErrorCode.CONFLICT
    assert collection_calls == 1
    assert workbook.read_bytes() == external_bytes
    refreshed = index_workbook(workbook)
    assert refreshed.changed is False
    with IndexStore(refreshed.index_path) as store:
        indexed = store.connection.execute(
            "SELECT value FROM cells WHERE sheet_id = 2 AND row = 2 AND col = 1"
        ).fetchone()
        assert indexed is not None
        assert indexed["value"] == 9999


def test_refresh_recalculated_clears_all_staleness_and_bumps_once(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "cross_sheet_model.xlsx")
    written = write_cells(workbook, (CellEdit.value("Inputs", "B2", 0.4),))

    refreshed = index_workbook(workbook, recalculated=True)

    assert refreshed.changed is True
    assert refreshed.reindexed_sheets == ()
    assert refreshed.generation == written.generation + 1
    with IndexStore(refreshed.index_path) as store:
        assert not store.is_stale("Calc", Rect(3, 3, 2, 2))
        assert not store.is_stale("Summary", Rect(3, 3, 3, 3))
        assert store.get_diagnostics(code="I_STALE").total == 0

    repeated = index_workbook(workbook, recalculated=True)
    assert repeated.changed is False
    assert repeated.generation == refreshed.generation


def test_external_recalculation_save_clears_staleness_for_changed_sheet(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "cross_sheet_model.xlsx")
    written = write_cells(
        workbook,
        (CellEdit.formula("Summary", "C2", "=Calc!D2*2"),),
    )
    root = _worksheet_root(workbook, "xl/worksheets/sheet3.xml")
    cell = _cells(root)["C2"]
    namespace = cell.tag[1:].split("}", 1)[0]
    value = etree.Element(f"{{{namespace}}}v")
    value.text = "600"
    cell.append(value)
    _rewrite_parts(workbook, {"xl/worksheets/sheet3.xml": etree.tostring(root)})

    refreshed = index_workbook(workbook)

    assert refreshed.changed is True
    assert refreshed.generation == written.generation + 1
    assert refreshed.reindexed_sheets == ("Summary",)
    with IndexStore(refreshed.index_path) as store:
        assert not store.is_stale("Summary", Rect(2, 2, 3, 3))
        row = store.connection.execute(
            "SELECT value FROM cells WHERE sheet_id = 3 AND row = 2 AND col = 3"
        ).fetchone()
        assert row is not None
        assert row["value"] == 600


def test_set_column_formula_requires_explicit_overwrite_and_fills_a1_pattern(
    tmp_path: Path,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    original = index_workbook(workbook)
    symbol = "col:Sales:0:linetotal"

    with pytest.raises(ExcelLSPError) as refused:
        set_column_formula(workbook, symbol, "=B2+C2")
    assert refused.value.code is ErrorCode.INVALID_VALUE
    assert index_workbook(workbook).generation == original.generation

    result = set_column_formula(workbook, symbol, "=B2+C2", overwrite=True)

    assert result.cells_written == 5
    assert result.formula_block.startswith("fblock:Sales:")
    cells = _cells(_worksheet_root(workbook))
    for row in range(2, 7):
        formula = child_by_local(cells[f"D{row}"], "f")
        assert formula is not None
        assert formula.text == f"B{row}+C{row}"
        assert child_by_local(cells[f"D{row}"], "v") is None


def test_set_column_formula_renders_r1c1_per_body_row_and_rejects_boundary_escape(
    tmp_path: Path,
) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    symbol = "col:Sales:0:linetotal"

    result = set_column_formula(
        workbook,
        symbol,
        "=RC[-2]*RC[-1]",
        overwrite=True,
    )

    assert result.cells_written == 5
    cells = _cells(_worksheet_root(workbook))
    for row in range(2, 7):
        formula = child_by_local(cells[f"D{row}"], "f")
        assert formula is not None
        assert formula.text == f"B{row}*C{row}"

    with pytest.raises(ExcelLSPError) as invalid:
        set_column_formula(
            workbook,
            symbol,
            "=R[-2]C",
            overwrite=True,
        )
    assert invalid.value.code is ErrorCode.INVALID_VALUE


def test_set_column_formula_rejects_unknown_symbol_before_writing(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "basic_single_table.xlsx")
    before = workbook.read_bytes()
    with pytest.raises(ExcelLSPError) as captured:
        set_column_formula(workbook, "col:Sales:99:missing", "=1")
    assert captured.value.code is ErrorCode.UNKNOWN_SYMBOL
    assert workbook.read_bytes() == before


def test_set_column_formula_is_not_limited_by_write_cells_batch_cap(tmp_path: Path) -> None:
    workbook = tmp_path / "large-column.xlsx"
    source = Workbook()
    sheet = source.active
    assert sheet is not None
    sheet.title = "Large"
    sheet.append(("Input", "Target"))
    for row in range(1, 502):
        sheet.append((row, row + 1))
    sheet.add_table(Table(displayName="LargeTable", ref="A1:B502"))
    source.save(workbook)
    source.close()

    result = set_column_formula(
        workbook,
        "col:Large:0:target",
        "=A2*2",
        overwrite=True,
    )

    assert result.cells_written == 501
    assert result.formula_block.startswith("fblock:Large:")
    cells = _cells(_worksheet_root(workbook))
    assert child_by_local(cells["B2"], "f").text == "A2*2"  # type: ignore[union-attr]
    assert child_by_local(cells["B502"], "f").text == "A502*2"  # type: ignore[union-attr]


def test_set_column_formula_excludes_the_listobject_totals_row(tmp_path: Path) -> None:
    workbook = _copy_fixture(tmp_path, "structured_table.xlsx")

    result = set_column_formula(
        workbook,
        "col:Structured:0:linetotal",
        "=B2*C2",
        overwrite=True,
    )

    assert result.cells_written == 4
    cells = _cells(_worksheet_root(workbook))
    for row in range(2, 6):
        formula = child_by_local(cells[f"D{row}"], "f")
        assert formula is not None
        assert formula.text == f"B{row}*C{row}"
    totals_formula = child_by_local(cells["D6"], "f")
    assert totals_formula is not None
    assert totals_formula.text == "SUBTOTAL(109,Table1[LineTotal])"


@pytest.mark.parametrize(
    ("fixture", "sheet", "protected_parts"),
    (
        (
            "macro_book.xlsm",
            "MacroModel",
            {"xl/vbaProject.bin"},
        ),
        (
            "chart_image.xlsx",
            "Dashboard",
            {
                "xl/charts/chart1.xml",
                "xl/drawings/drawing1.xml",
                "xl/drawings/_rels/drawing1.xml.rels",
                "xl/media/image1.png",
            },
        ),
    ),
)
def test_f16_f21_edits_preserve_every_untouched_part_byte_for_byte(
    tmp_path: Path,
    fixture: str,
    sheet: str,
    protected_parts: set[str],
) -> None:
    workbook = _copy_fixture(tmp_path, fixture)
    before = _parts(workbook)

    result = write_cells(workbook, (CellEdit.value(sheet, "A2", 99),))

    after = _parts(workbook)
    assert protected_parts <= before.keys() & after.keys()
    assert set(result.patch.modified_parts) == {"xl/worksheets/sheet1.xml"}
    for part, payload in before.items():
        if part not in result.patch.modified_parts:
            assert after[part] == payload


@pytest.mark.parametrize("fixture_id", ("F16", "F21"))
def test_committed_part_diff_evidence_matches_fresh_render(fixture_id: str) -> None:
    root = Path(__file__).parents[2]
    completed = subprocess.run(
        [sys.executable, "tests/fixtures/render_part_diff.py", fixture_id],
        cwd=root,
        check=True,
        capture_output=True,
        text=True,
    )
    actual = json.loads(completed.stdout)
    evidence = root / "docs" / "evidence" / f"part-diff-{fixture_id.casefold()}.json"
    expected = json.loads(evidence.read_text(encoding="utf-8"))

    assert actual == expected
    assert actual["untouchedPartsByteIdentical"] is True
    assert actual["protectedPartsByteIdentical"] is True
