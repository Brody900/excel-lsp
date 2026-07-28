"""Desktop-Excel smoke tests excluded from automated CI."""

from __future__ import annotations

import gc
import json
import shutil
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

import lxml.etree as etree
import pytest

from excel_lsp.core.edit import CellEdit, write_cells
from excel_lsp.core.index.lifecycle import index_workbook
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import Rect
from excel_lsp.core.parse._xml import child_by_local, local_name, parse_xml

pytestmark = pytest.mark.live

FIXTURES = Path(__file__).parents[1] / "fixtures" / "generated"


def _rewrite_parts(path: Path, replacements: dict[str, bytes]) -> None:
    temporary = path.with_suffix(".rewrite")
    with (
        ZipFile(path, "r") as source,
        ZipFile(temporary, "w", compression=ZIP_DEFLATED) as target,
    ):
        for info in source.infolist():
            payload = replacements.get(
                info.filename,
                source.read(info) if not info.is_dir() else b"",
            )
            target.writestr(info, payload)
    temporary.replace(path)


def _inject_schema_order_adversary(path: Path) -> None:
    with ZipFile(path) as archive:
        workbook_root = parse_xml(archive.read("xl/workbook.xml"))
        sheet_root = parse_xml(archive.read("xl/worksheets/sheet1.xml"))

    calc = child_by_local(workbook_root, "calcPr")
    assert calc is not None
    workbook_root.remove(calc)
    workbook_namespace = workbook_root.tag[1:].split("}", 1)[0]
    etree.SubElement(
        workbook_root,
        f"{{{workbook_namespace}}}fileRecoveryPr",
        autoRecover="1",
    )

    input_cell = next(
        element
        for element in sheet_root.iter()
        if isinstance(element.tag, str)
        and local_name(element.tag) == "c"
        and element.get("r") == "B2"
    )
    cell_namespace = input_cell.tag[1:].split("}", 1)[0]
    extension_list = etree.SubElement(input_cell, f"{{{cell_namespace}}}extLst")
    etree.SubElement(
        extension_list,
        f"{{{cell_namespace}}}ext",
        uri="{25C4219D-8188-4D27-B50F-8B7A0A88E7D6}",
    )
    _rewrite_parts(
        path,
        {
            "xl/workbook.xml": etree.tostring(workbook_root),
            "xl/worksheets/sheet1.xml": etree.tostring(sheet_root),
        },
    )


def test_surgical_edit_round_trips_through_desktop_excel(tmp_path: Path) -> None:
    win32com = pytest.importorskip("win32com.client")
    pythoncom = pytest.importorskip("pythoncom")
    workbook = tmp_path / "p6-live-round-trip.xlsx"
    shutil.copyfile(FIXTURES / "cross_sheet_model.xlsx", workbook)
    _inject_schema_order_adversary(workbook)

    written = write_cells(workbook, (CellEdit.value("Inputs", "B2", 0.2),))
    with ZipFile(workbook) as archive:
        workbook_root = parse_xml(archive.read("xl/workbook.xml"))
        sheet_root = parse_xml(archive.read("xl/worksheets/sheet1.xml"))
    workbook_names = [local_name(child.tag) for child in workbook_root]
    assert workbook_names.index("calcPr") < workbook_names.index("fileRecoveryPr")
    input_cell = next(
        element
        for element in sheet_root.iter()
        if isinstance(element.tag, str)
        and local_name(element.tag) == "c"
        and element.get("r") == "B2"
    )
    assert [local_name(child.tag) for child in input_cell] == ["v", "extLst"]
    with IndexStore(workbook.with_name(f"{workbook.name}.xlsp.db")) as store:
        assert store.is_stale("Summary", Rect(10, 10, 3, 3))

    application: Any = None
    opened: Any = None
    pythoncom.CoInitialize()
    try:
        application = win32com.DispatchEx("Excel.Application")
        application.Visible = False
        application.DisplayAlerts = False
        application.AskToUpdateLinks = False
        opened = application.Workbooks.Open(
            str(workbook),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            CorruptLoad=0,
        )
        assert float(opened.Worksheets("Inputs").Range("B2").Value) == pytest.approx(0.2)
        application.CalculateFullRebuild()
        total_after_tax = float(opened.Worksheets("Summary").Range("C10").Value)
        assert total_after_tax == pytest.approx(2232.48)
        excel_version = str(application.Version)
        excel_build = str(application.Build)
        opened.Save()
        opened.Close(SaveChanges=False)
        opened = None
    finally:
        try:
            if opened is not None:
                opened.Close(SaveChanges=False)
            if application is not None:
                application.Quit()
        finally:
            opened = None
            application = None
            gc.collect()
            pythoncom.CoUninitialize()

    refreshed = index_workbook(workbook, recalculated=True)
    assert refreshed.generation == written.generation + 1
    with IndexStore(refreshed.index_path) as store:
        assert not store.is_stale("Summary", Rect(10, 10, 3, 3))
        stored = store.connection.execute(
            "SELECT value FROM cells WHERE sheet_id = 3 AND row = 10 AND col = 3"
        ).fetchone()
        assert stored is not None
        assert float(stored["value"]) == pytest.approx(2232.48)

    artifact = {
        "artifactVersion": 1,
        "fixture": "F03 cross_sheet_model.xlsx",
        "edit": {"sheet": "Inputs", "ref": "B2", "value": 0.2},
        "excel": {"version": excel_version, "build": excel_build},
        "normalOpen": True,
        "fullCalculation": True,
        "savedAndClosed": True,
        "expectedSummaryC10": 2232.48,
        "observedSummaryC10": total_after_tax,
        "stalenessBeforeExcelSave": True,
        "stalenessAfterRecalculatedRefresh": False,
        "schemaOrderAdversary": {
            "cellExtLstPreservedAfterValue": True,
            "calcPrCreatedBeforeFileRecoveryPr": True,
        },
    }
    evidence = Path(__file__).parents[2] / "docs" / "evidence" / "live-excel" / "p6-smoke.json"
    assert artifact == json.loads(evidence.read_text(encoding="utf-8"))
    print(json.dumps(artifact, sort_keys=True))
