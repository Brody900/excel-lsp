from __future__ import annotations

import runpy
from collections.abc import Callable
from pathlib import Path
from typing import cast

from excel_lsp.core.values import JsonScalar

CanonicalCell = tuple[str, str, JsonScalar, str | None]
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


def test_openpyxl_dual_load_observes_f01_formula_caches(tmp_path: Path) -> None:
    canonical = openpyxl_canonical_cells(generate_all(tmp_path)["F01"])
    by_ref = {(sheet, ref): (value, formula) for sheet, ref, value, formula in canonical}

    assert by_ref[("Sales", "D2")] == (7, "=B2*C2")
    assert by_ref[("Sales", "D3")] == (6.25, "=B3*C3")
    assert by_ref[("Sales", "D6")] == (27, "=B6*C6")


def test_openpyxl_315_read_only_probe_matches_recorded_behavior(tmp_path: Path) -> None:
    result = probe_read_only(generate_all(tmp_path)["F07"])

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


def test_f01_matches_production_ooxml_parser(tmp_path: Path) -> None:
    path = generate_all(tmp_path)["F01"]
    assert ooxml_canonical_cells(path) == openpyxl_canonical_cells(path)


def test_f07_matches_production_ooxml_parser(tmp_path: Path) -> None:
    path = generate_all(tmp_path)["F07"]
    assert ooxml_canonical_cells(path) == openpyxl_canonical_cells(path)
