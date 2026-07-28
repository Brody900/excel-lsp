from __future__ import annotations

import runpy
from collections.abc import Callable
from pathlib import Path
from typing import cast

import pytest

from excel_lsp.core.values import JsonScalar

CanonicalCell = tuple[str, str, JsonScalar, str | None]
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
    "F16",
    "F18",
    "F19",
    "F20",
    "F21",
}
GenerateAll = Callable[[Path], dict[str, Path]]
CanonicalReader = Callable[[Path], tuple[CanonicalCell, ...]]
Probe = Callable[[Path], dict[str, object]]

generate_all = cast(
    GenerateAll,
    runpy.run_path(str(Path(__file__).parents[1] / "fixtures" / "generate.py"))["generate_all"],
)
_oracle_exports = runpy.run_path(str(Path(__file__).with_name("oracle.py")))
openpyxl_canonical_cells = cast(CanonicalReader, _oracle_exports["openpyxl_canonical_cells"])
ooxml_canonical_cells = cast(CanonicalReader, _oracle_exports["ooxml_canonical_cells"])
probe_read_only = cast(
    Probe,
    runpy.run_path(str(Path(__file__).with_name("probe_openpyxl.py")))["probe_read_only"],
)


@pytest.fixture(scope="module")
def generated_paths(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    """Share the complete corpus so the 50k-row oracle fixture is authored once."""
    return generate_all(tmp_path_factory.mktemp("oracle-corpus"))


def test_openpyxl_dual_load_observes_f01_formula_caches(
    generated_paths: dict[str, Path],
) -> None:
    canonical = openpyxl_canonical_cells(generated_paths["F01"])
    by_ref = {(sheet, ref): (value, formula) for sheet, ref, value, formula in canonical}

    assert by_ref[("Sales", "D2")] == (7, "=B2*C2")
    assert by_ref[("Sales", "D3")] == (6.25, "=B3*C3")
    assert by_ref[("Sales", "D6")] == (27, "=B6*C6")


def test_openpyxl_dual_load_observes_f03_cross_sheet_caches(
    generated_paths: dict[str, Path],
) -> None:
    canonical = openpyxl_canonical_cells(generated_paths["F03"])
    by_ref = {(sheet, ref): (value, formula) for sheet, ref, value, formula in canonical}

    assert by_ref[("Calc", "B3")] == (1100, "=B2*(1+Inputs!$B$2)")
    assert by_ref[("Calc", "D6")] == (439.23, "=(B6-C6)*(1-Inputs!$B$5)")
    assert by_ref[("Summary", "C8")] == (1464.1, "=Calc!B6")
    assert by_ref[("Summary", "C10")] == (1831.53, "=SUM(Calc!D2:D6)")


def test_openpyxl_315_read_only_probe_matches_recorded_behavior(
    generated_paths: dict[str, Path],
) -> None:
    result = probe_read_only(generated_paths["F07"])

    assert result == {
        "openpyxl_version": "3.1.5",
        "worksheet_type": "ReadOnlyWorksheet",
        "shared_formula_followers": {
            "C3": "=A3*B3",
            "C11": "=A11*B11",
            "C14": "=A14*B14",
            "C21": "=A21*B21",
        },
        "formula_caches": {
            "C2": 4,
            "C3": 10,
            "C11": 30,
            "C12": 15,
            "C13": 60,
            "C14": 78,
            "C21": 60,
        },
        "read_only_tables": {
            "available": False,
            "error_type": "AttributeError",
        },
        "read_only_merged_cells": {
            "available": False,
            "error_type": "AttributeError",
        },
        "normal_mode_control": {
            "table_names": ["FormulaBlocksTable"],
            "merged_ranges": ["E1:F1"],
        },
    }


def test_openpyxl_dual_load_observes_f19_modern_formulas_and_caches(
    generated_paths: dict[str, Path],
) -> None:
    canonical = openpyxl_canonical_cells(generated_paths["F19"])
    by_ref = {(sheet, ref): (value, formula) for sheet, ref, value, formula in canonical}

    assert by_ref[("Modern", "A1")] == (
        20,
        "=_xlfn._xlws.FILTER(I2:I4,I2:I4>=20)",
    )
    assert by_ref[("Modern", "A2")] == (30, None)
    assert by_ref[("Modern", "B1")] == (50, "=SUM(A1#)")
    assert by_ref[("Modern", "C1")] == (50, "=SUM(FilteredValues#)")
    assert by_ref[("Modern", "D1")] == (
        31,
        "=_xlfn.LET(_xlpm.rate,I2,_xlpm.bonus,1,_xlpm.rate*3+_xlpm.bonus)",
    )
    assert by_ref[("Modern", "E1")] == (40, "=DoubleIt(I3)")
    assert by_ref[("Modern", "F1")] == (
        20,
        '=_xlfn.XLOOKUP("beta",H2:H4,I2:I4,"missing")',
    )
    assert by_ref[("Modern", "G2")] == (10, "=@I2:I4")


def test_openpyxl_observes_p4_names_structured_cycle_and_3d_caches(
    generated_paths: dict[str, Path],
) -> None:
    expected = {
        "F04": {
            ("Inputs", "B4"): (110, "=BaseAmount*(1+GlobalRate)"),
            ("Calc", "B4"): (105, "=BaseAmount*(1+ScopedRate)"),
        },
        "F05": {
            ("Structured", "D2"): (7, "=[@Qty]*[@Price]"),
            ("Structured", "D6"): (54.25, "=SUBTOTAL(109,Table1[LineTotal])"),
            ("Structured", "F2"): (54.25, "=SUM(Table1[LineTotal])"),
        },
        "F09a": {
            ("Circular", "B2"): (0, "=B3+1"),
            ("Circular", "B3"): (0, "=B2+1"),
        },
        "F09b": {
            ("RunningTotal", "B2"): (0, None),
            ("RunningTotal", "B3"): (0, "=SUM($B$2:B2)"),
            ("RunningTotal", "B50002"): (0, "=SUM($B$2:B50001)"),
        },
        "F15": {
            ("Summary", "B2"): (60, "=SUM(Jan:Mar!B2)"),
        },
    }
    for fixture_id, expected_cells in expected.items():
        canonical = openpyxl_canonical_cells(generated_paths[fixture_id])
        by_ref = {(sheet, ref): (value, formula) for sheet, ref, value, formula in canonical}
        assert {ref: by_ref[ref] for ref in expected_cells} == expected_cells


def test_openpyxl_observes_p5_error_dynamic_external_and_volatile_caches(
    generated_paths: dict[str, Path],
) -> None:
    expected = {
        "F08": {
            ("Errors", "B2"): ("#REF!", "=NA()"),
            ("Errors", "B10"): ("#BLOCKED!", "=NA()"),
            ("Errors", "B11"): ("#FIELD!", "=NA()"),
        },
        "F10": {("External", "A2"): (0, "=[1]Data!A1")},
        "F11": {
            ("DynamicRefs", "B2"): (10, '=INDIRECT("A2")'),
            ("DynamicRefs", "C2"): (20, "=OFFSET(A2,1,0)"),
        },
        "F18": {
            ("Volatile", "B2"): (45292.5, "=NOW()"),
            ("Volatile", "B3"): (0.25, "=RAND()"),
        },
    }
    for fixture_id, expected_cells in expected.items():
        canonical = openpyxl_canonical_cells(generated_paths[fixture_id])
        by_ref = {(sheet, ref): (value, formula) for sheet, ref, value, formula in canonical}
        assert {ref: by_ref[ref] for ref in expected_cells} == expected_cells


def test_every_emitted_fixture_matches_production_ooxml_parser(
    generated_paths: dict[str, Path],
) -> None:
    assert generated_paths.keys() == EXPECTED_FIXTURE_IDS
    for fixture_id, path in generated_paths.items():
        assert ooxml_canonical_cells(path) == openpyxl_canonical_cells(path), fixture_id
