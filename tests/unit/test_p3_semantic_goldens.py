"""Deterministic fixture-level snapshots for P3 formula semantics."""

from __future__ import annotations

import json
import os
import runpy
from collections.abc import Callable
from pathlib import Path
from typing import cast

import pytest

from excel_lsp.core.index import IndexStore, index_workbook

GenerateAll = Callable[[Path], dict[str, Path]]
generate_all = cast(
    GenerateAll,
    runpy.run_path(str(Path(__file__).parents[1] / "fixtures" / "generate.py"))["generate_all"],
)

_FIXTURE_IDS = ("F07", "F19")
_UPDATE_ENV = "EXCEL_LSP_UPDATE_P3_GOLDENS"


@pytest.fixture(scope="module")
def p3_golden_fixtures(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    return generate_all(tmp_path_factory.mktemp("p3-golden-fixtures"))


@pytest.mark.parametrize("fixture_id", _FIXTURE_IDS)
def test_p3_semantic_snapshot_matches_or_updates_golden(
    tmp_path: Path,
    p3_golden_fixtures: dict[str, Path],
    fixture_id: str,
) -> None:
    """Compare exact bytes; set EXCEL_LSP_UPDATE_P3_GOLDENS=1 to refresh."""
    update = index_workbook(
        p3_golden_fixtures[fixture_id],
        index_dir=tmp_path / "indexes",
    )
    with IndexStore(update.index_path) as store:
        exported = store.canonical_export()

    snapshot = {
        "fixture": fixture_id,
        "formulaCells": tuple(
            (row[0], row[3], row[6], row[8], row[9])
            for row in exported["cells"]
            if row[6] is not None
        ),
        "fblocks": exported["fblocks"],
        "edges": exported["edges"],
        "diagnostics": exported["diagnostics"],
    }
    serialized = json.dumps(snapshot, ensure_ascii=False, indent=2) + "\n"
    golden_path = (
        Path(__file__).parents[1] / "golden" / f"{fixture_id.lower()}-formula-semantics.json"
    )
    if os.environ.get(_UPDATE_ENV) == "1":
        golden_path.write_text(serialized, encoding="utf-8", newline="\n")

    assert golden_path.read_bytes() == serialized.encode("utf-8")
