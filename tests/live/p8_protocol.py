"""Execute exact COM assertions for the required P8 desktop-Excel protocol."""

from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

from excel_lsp.server.models import WriteCellInput
from excel_lsp.server.service import ToolService

ROOT = Path(__file__).parents[2]
LIVE = ROOT / "tests" / "fixtures" / "live"
GENERATED = ROOT / "tests" / "fixtures" / "generated"
EVIDENCE = ROOT / "docs" / "evidence" / "live-excel"
WORK = ROOT / ".tmp-live-excel"
L1 = LIVE / "L1-model.xlsx"
L2 = LIVE / "L2-table.xlsx"
L3 = LIVE / "L3-broken.xlsx"
F16 = WORK / "F16-roundtrip.xlsm"
F21 = WORK / "F21-roundtrip.xlsx"


def _write(name: str, payload: dict[str, Any]) -> None:
    EVIDENCE.mkdir(parents=True, exist_ok=True)
    (EVIDENCE / name).write_text(
        json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8", newline="\n"
    )


def _application() -> Any:
    application = win32com.client.GetActiveObject("Excel.Application")
    application.Visible = True
    application.DisplayAlerts = True
    return application


def _workbook(application: Any, name: str) -> Any:
    for workbook in application.Workbooks:
        if str(workbook.Name).casefold() == name.casefold():
            return workbook
    raise AssertionError(f"Excel workbook is not open: {name}")


def assert_authored() -> dict[str, Any]:
    application = _application()
    l1 = _workbook(application, L1.name)
    l2 = _workbook(application, L2.name)
    l3 = _workbook(application, L3.name)

    assert [str(sheet.Name) for sheet in l1.Worksheets] == ["Inputs", "Model"]
    assert str(l1.Names("PrimaryInput").RefersTo) == "=Inputs!$B$2"
    model = l1.Worksheets("Model")
    assert str(model.Range("A2").Formula) == "=Inputs!B2"
    assert str(model.Range("B2").Formula) == "=A2*2"
    assert str(model.Range("C2").Formula) == "=A2+B2"
    assert str(model.Range("A4").Formula) == "=Inputs!B4"

    table = l2.Worksheets(1).ListObjects("Table1")
    assert bool(table.ShowTotals)
    assert str(table.Range.Address) == "$A$1:$D$4"
    assert "SUBTOTAL" in str(table.TotalsRowRange.Cells(1, 4).Formula).upper()
    assert float(table.TotalsRowRange.Cells(1, 4).Value) == 31.0

    broken = l3.Worksheets(1)
    assert [str(broken.Range(f"B{row}").Text) for row in range(2, 7)] == ["#REF!"] * 5
    assert [str(broken.Range(f"B{row}").Formula) for row in range(2, 7)] == ["=#REF!+1"] * 5
    assert str(broken.Range("C4").Formula) == "=A4*99"
    assert float(broken.Range("C4").Value) == 297.0
    assert str(broken.Range("C3").Formula) == "=A3*3"
    assert str(broken.Range("C5").Formula) == "=A5*3"

    payload = {
        "artifactVersion": 1,
        "excel": {"version": str(application.Version), "build": str(application.Build)},
        "L1": {
            "sheets": ["Inputs", "Model"],
            "name": {"PrimaryInput": "=Inputs!$B$2"},
            "fillDown": {"A2": "=Inputs!B2", "B2": "=A2*2", "C2": "=A2+B2"},
        },
        "L2": {
            "table": "Table1",
            "range": "$A$1:$D$4",
            "totalsRow": True,
            "total": 31,
        },
        "L3": {
            "brokenRefs": [f"Sheet1!B{row}" for row in range(2, 7)],
            "brokenValue": "#REF!",
            "tamperedCell": "Sheet1!C4",
            "tamperedFormula": "=A4*99",
            "tamperedValue": 297,
        },
    }
    _write("authoring.json", payload)
    return payload


def inspect_product() -> dict[str, Any]:
    service = ToolService()
    l1_map = service.open_workbook(str(L1))
    l2_map = service.open_workbook(str(L2))
    l3_diagnostics = service.get_diagnostics(str(L3))

    l2_regions = l2_map["sheetList"][0]["regions"]
    assert any(
        region["kind"] == "table" and region["range"] == "A1:D4" and region["conf"] == 1.0
        for region in l2_regions
    )
    assert l3_diagnostics["counts"]["code"] == {
        "E_ERRVAL": 5,
        "W_INCONSISTENT_FORMULA": 1,
    }
    error_refs = {
        item["ref"] for item in l3_diagnostics["diagnostics"] if item["code"] == "E_ERRVAL"
    }
    assert error_refs == {f"cell:Sheet1!B{row}" for row in range(2, 7)}
    assert {
        item["ref"]
        for item in l3_diagnostics["diagnostics"]
        if item["code"] == "W_INCONSISTENT_FORMULA"
    } == {"cell:Sheet1!C4"}

    _write("01-l1-map.json", l1_map)
    _write("01-l2-map.json", l2_map)
    _write("02-l3-diagnostics.json", l3_diagnostics)
    return {
        "artifactVersion": 1,
        "L1": {"sheets": l1_map["sheets"], "diagCounts": l1_map["diagCounts"]},
        "L2": {"tableRange": "A1:D4", "confidence": 1.0},
        "L3": l3_diagnostics["counts"],
    }


def prepare_writes() -> dict[str, Any]:
    application = _application()
    open_names = {str(workbook.Name).casefold() for workbook in application.Workbooks}
    assert L1.name.casefold() not in open_names, "close L1 before surgical writes"
    WORK.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(GENERATED / "macro_book.xlsm", F16)
    shutil.copyfile(GENERATED / "chart_image.xlsx", F21)

    service = ToolService()
    l1_write = service.write_cells(str(L1), [WriteCellInput(ref="Inputs!B2", value=15)])
    l1_formula = service.set_column_formula(
        str(L1), "col:Model:0:double", "=RC[-1]*4", overwrite=True
    )
    f16_write = service.write_cells(str(F16), [WriteCellInput(ref="MacroModel!A2", value=21)])
    f16_formula = service.set_column_formula(
        str(F16), "col:MacroModel:0:doubled", "=RC[-1]*2", overwrite=True
    )
    f21_write = service.write_cells(str(F21), [WriteCellInput(ref="Dashboard!B2", value=99)])
    for result in (l1_write, f16_write, f21_write):
        assert all(bool(item["ok"]) for item in result["results"])
    assert int(l1_formula["cellsWritten"]) == 3
    assert int(f16_formula["cellsWritten"]) == 1

    payload = {
        "artifactVersion": 1,
        "L1": {"write_cells": l1_write, "set_column_formula": l1_formula},
        "F16": {"write_cells": f16_write, "set_column_formula": f16_formula},
        "F21": {"write_cells": f21_write},
    }
    _write("product-writes.json", payload)
    return payload


def open_roundtrip() -> dict[str, Any]:
    application = _application()
    try:
        l1 = _workbook(application, L1.name)
    except AssertionError:
        l1 = application.Workbooks.Open(
            str(L1), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, CorruptLoad=0
        )
    try:
        f16 = _workbook(application, F16.name)
    except AssertionError:
        f16 = application.Workbooks.Open(
            str(F16), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, CorruptLoad=0
        )
    application.CalculateFullRebuild()

    l1_model = l1.Worksheets("Model")
    assert float(l1.Worksheets("Inputs").Range("B2").Value) == 15.0
    assert [float(l1_model.Range(f"B{row}").Value) for row in range(2, 5)] == [
        60.0,
        80.0,
        120.0,
    ]
    assert [float(l1_model.Range(f"C{row}").Value) for row in range(2, 5)] == [
        75.0,
        100.0,
        150.0,
    ]
    macro_sheet = f16.Worksheets("MacroModel")
    assert float(macro_sheet.Range("A2").Value) == 21.0
    assert float(macro_sheet.Range("B2").Value) == 42.0
    f16.Activate()
    macro_sheet.Activate()
    application.Run(f"'{f16.Name}'!Stamp")
    assert float(macro_sheet.Range("Z1").Value) == 42.0
    l1.Save()
    f16.Save()

    service = ToolService()
    l1_refresh = service.refresh(str(L1), recalculated=True)
    f16_refresh = service.refresh(str(F16), recalculated=True)
    refusal = service.write_cells(str(L1), [WriteCellInput(ref="Inputs!B3", value=999)])
    assert refusal["results"][0]["ok"] is False
    assert refusal["results"][0]["error"]["code"] in {"E_OPEN_IN_EXCEL", "E_LOCKED"}

    traces = {
        "Model!A2": service.trace_precedents(str(L1), "Model!A2", 1, 20),
        "Model!B2": service.trace_precedents(str(L1), "Model!B2", 1, 20),
        "Model!C2": service.trace_precedents(str(L1), "Model!C2", 1, 20),
    }
    expected = {
        "Model!A2": {"Inputs!B2:B4"},
        "Model!B2": {"Model!A2:A4"},
        "Model!C2": {"Model!A2:A4", "Model!B2:B4"},
    }
    for ref, trace in traces.items():
        tree = trace["tree"]
        assert tree["symbol"] == f"cell:{ref}"
        observed = {str(node["ref"]) for node in tree["children"]}
        assert observed == expected[ref]

    f16.Activate()
    macro_sheet.Activate()
    macro_sheet.Range("Z1").Select()
    payload = {
        "artifactVersion": 1,
        "excel": {"version": str(application.Version), "build": str(application.Build)},
        "normalOpen": {"L1": True, "F16": True},
        "repairDialog": {"L1": False, "F16": False},
        "recalculation": {
            "L1": {"ModelB2_B4": [60, 80, 120], "ModelC2_C4": [75, 100, 150]},
            "F16": {"MacroModelA2": 21, "MacroModelB2": 42},
        },
        "macro": {"name": "Stamp", "ran": True, "MacroModelZ1": 42},
        "refresh": {"L1": l1_refresh, "F16": f16_refresh},
        "writeRefusal": refusal,
        "traces": traces,
    }
    _write("roundtrip.json", payload)
    _write("write-refusal.json", {"artifactVersion": 1, "result": refusal})
    return payload


def open_chart() -> dict[str, Any]:
    application = _application()
    try:
        l1 = _workbook(application, L1.name)
        l1.Close(SaveChanges=True)
    except AssertionError:
        pass
    try:
        f16 = _workbook(application, F16.name)
        f16.Close(SaveChanges=True)
    except AssertionError:
        pass
    try:
        workbook = _workbook(application, F21.name)
    except AssertionError:
        workbook = application.Workbooks.Open(
            str(F21), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True, CorruptLoad=0
        )
    sheet = workbook.Worksheets("Dashboard")
    assert float(sheet.Range("B2").Value) == 99.0
    chart_objects = sheet.ChartObjects()
    assert int(chart_objects.Count) == 1
    chart_name = str(chart_objects(1).Name)
    shapes = [sheet.Shapes(index) for index in range(1, int(sheet.Shapes.Count) + 1)]
    shape_names = [str(shape.Name) for shape in shapes]
    assert len(shape_names) >= 2
    assert chart_name in shape_names
    assert "Image 2" in shape_names
    shape_geometry = {
        str(shape.Name): {
            "type": int(shape.Type),
            "visible": int(shape.Visible) != 0,
            "left": round(float(shape.Left), 3),
            "top": round(float(shape.Top), 3),
            "width": round(float(shape.Width), 3),
            "height": round(float(shape.Height), 3),
        }
        for shape in shapes
    }
    for name in (chart_name, "Image 2"):
        geometry = shape_geometry[name]
        assert geometry["visible"] is True
        assert geometry["width"] > 0
        assert geometry["height"] > 0
    workbook.Activate()
    sheet.Activate()
    sheet.Shapes("Image 2").Select()
    selected_shape = str(application.Selection.Name)
    assert selected_shape == "Image 2"
    payload = {
        "artifactVersion": 1,
        "normalOpen": True,
        "repairDialog": False,
        "editedCell": {"ref": "Dashboard!B2", "value": 99},
        "chartObjects": int(chart_objects.Count),
        "shapes": shape_names,
        "shapeGeometry": shape_geometry,
        "chartVisible": shape_geometry[chart_name]["visible"],
        "imageVisible": shape_geometry["Image 2"]["visible"],
        "selectedShape": selected_shape,
    }
    _write("chart-preservation.json", payload)
    return payload


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "mode",
        choices=(
            "assert-authored",
            "inspect-product",
            "prepare-writes",
            "open-roundtrip",
            "open-chart",
        ),
    )
    mode = parser.parse_args().mode
    pythoncom.CoInitialize()
    try:
        payload = {
            "assert-authored": assert_authored,
            "inspect-product": inspect_product,
            "prepare-writes": prepare_writes,
            "open-roundtrip": open_roundtrip,
            "open-chart": open_chart,
        }[mode]()
        print(json.dumps(payload, sort_keys=True))
    finally:
        pythoncom.CoUninitialize()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
