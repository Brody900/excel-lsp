"""Fixture-level region persistence, configuration, and map budget evidence."""

from __future__ import annotations

import json
import runpy
from collections.abc import Callable
from datetime import datetime
from importlib.metadata import version as distribution_version
from pathlib import Path
from typing import cast

import pytest
import tiktoken

from excel_lsp.core.index import IndexStore, index_workbook
from excel_lsp.core.parse.coordinates import make_cell_ref
from excel_lsp.core.workbook_map import build_workbook_map, serialize_workbook_map

GenerateAll = Callable[[Path], dict[str, Path]]

generate_all = cast(
    GenerateAll,
    runpy.run_path(str(Path(__file__).parents[1] / "fixtures" / "generate.py"))["generate_all"],
)


@pytest.fixture(scope="module")
def p2_fixtures(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    return generate_all(tmp_path_factory.mktemp("p2-fixtures"))


def test_native_table_region_and_profiles_persist_exactly(
    tmp_path: Path,
    p2_fixtures: dict[str, Path],
) -> None:
    update = index_workbook(p2_fixtures["F01"], index_dir=tmp_path / "indexes")

    with IndexStore(update.index_path) as store:
        region = store.connection.execute(
            """
            SELECT r.n, r.row_min, r.row_max, r.col_min, r.col_max,
                   r.header_rows, r.kind, r.list_object_name, r.confidence
            FROM regions AS r JOIN sheets AS s ON s.id = r.sheet_id
            WHERE s.name = 'Sales'
            """
        ).fetchone()
        columns = store.connection.execute(
            """
            SELECT c.header, c.norm_header, c.dtype, c.nonnull, c.distinct_est
            FROM columns AS c
            JOIN regions AS r ON r.id = c.region_id
            JOIN sheets AS s ON s.id = r.sheet_id
            WHERE s.name = 'Sales' ORDER BY c.idx
            """
        ).fetchall()

    assert tuple(region) == (0, 1, 6, 1, 4, 1, "table", "SalesTable", 1.0)
    assert [tuple(row) for row in columns] == [
        ("Item", "item", "str", 5, 5),
        ("Quantity", "quantity", "int", 5, 5),
        ("UnitPrice", "unitprice", "float", 5, 5),
        ("LineTotal", "linetotal", "float", 5, 5),
    ]


def test_gap_tolerance_is_configurable_and_invalidates_freshness(
    tmp_path: Path,
    p2_fixtures: dict[str, Path],
) -> None:
    workbook = p2_fixtures["F02"]
    index_dir = tmp_path / "indexes"
    default = index_workbook(workbook, index_dir=index_dir)
    with IndexStore(default.index_path) as store:
        assert _region_ranges(store, "Islands") == (
            "A1:C6",
            "F2:H6",
            "B10:D13",
        )

    strict = index_workbook(workbook, index_dir=index_dir, gap_tol=0)

    assert strict.changed is True
    assert strict.generation == default.generation + 1
    with IndexStore(strict.index_path) as store:
        strict_indexed_at = store.get_meta("indexed_at")
        assert store.get_meta("region_gap_tol") == "0"
        assert _region_ranges(store, "Islands") == (
            "A1:C3",
            "F2:H6",
            "A5:C6",
            "B10:D13",
        )

    build_workbook_map(workbook, index_dir=index_dir)
    with IndexStore(strict.index_path) as store:
        assert store.generation == strict.generation
        assert store.get_meta("indexed_at") == strict_indexed_at
        assert store.get_meta("region_gap_tol") == "0"
        assert _region_ranges(store, "Islands") == (
            "A1:C3",
            "F2:H6",
            "A5:C6",
            "B10:D13",
        )

    reset = index_workbook(workbook, index_dir=index_dir, gap_tol=1)
    assert reset.changed is True
    assert reset.generation == strict.generation + 1
    with IndexStore(reset.index_path) as store:
        assert store.get_meta("region_gap_tol") == "1"
        assert _region_ranges(store, "Islands") == (
            "A1:C6",
            "F2:H6",
            "B10:D13",
        )

    noop = index_workbook(workbook, index_dir=index_dir)
    assert noop.changed is False
    assert noop.generation == reset.generation


def test_merged_headers_mixed_types_and_sparse_sheets_match_fixture_contracts(
    tmp_path: Path,
    p2_fixtures: dict[str, Path],
) -> None:
    expected = {
        "F12": (
            "MergedHeaders",
            (
                ("Region", "str"),
                ("Revenue / Q1", "float"),
                ("Revenue / Q2", "float"),
                ("Revenue / Q3", "float"),
                ("Units / Actual", "int"),
                ("Units / Target", "int"),
            ),
        ),
        "F13": (
            "MixedTypes",
            (
                ("RecordID", "int"),
                ("PostingDate", "date"),
                ("Amount", "float"),
                ("MarginPct", "float"),
                ("AccountCode", "str"),
                ("Approved", "bool"),
                ("MixedSample", "mixed"),
            ),
        ),
    }
    for fixture_id, (sheet_name, expected_columns) in expected.items():
        update = index_workbook(
            p2_fixtures[fixture_id],
            index_dir=tmp_path / f"index-{fixture_id}",
        )
        with IndexStore(update.index_path) as store:
            rows = store.connection.execute(
                """
                SELECT c.header, c.dtype
                FROM columns AS c
                JOIN regions AS r ON r.id = c.region_id
                JOIN sheets AS s ON s.id = r.sheet_id
                WHERE s.name = ? ORDER BY r.n, c.idx
                """,
                (sheet_name,),
            ).fetchall()
        assert tuple(tuple(row) for row in rows) == expected_columns

    sparse = index_workbook(p2_fixtures["F14"], index_dir=tmp_path / "index-F14")
    with IndexStore(sparse.index_path) as store:
        sheets = store.connection.execute(
            """
            SELECT s.name, s.max_row, s.max_col, COUNT(r.id)
            FROM sheets AS s LEFT JOIN regions AS r ON r.sheet_id = s.id
            GROUP BY s.id ORDER BY s.id
            """
        ).fetchall()
        assert [tuple(row) for row in sheets] == [
            ("EmptyBefore", 0, 0, 0),
            ("LoneCells", 100, 24, 2),
            ("EmptyAfter", 0, 0, 0),
        ]
        assert _region_ranges(store, "LoneCells") == ("B2", "X100")


@pytest.mark.parametrize("fixture_id", ["F03", "F20"])
def test_workbook_map_matches_golden_and_budget(
    tmp_path: Path,
    p2_fixtures: dict[str, Path],
    fixture_id: str,
) -> None:
    result = build_workbook_map(
        p2_fixtures[fixture_id],
        index_dir=tmp_path / f"index-{fixture_id}",
    )
    serialized = serialize_workbook_map(result)
    indexed_at = result.get("indexed_at")
    assert isinstance(indexed_at, str) and indexed_at
    assert datetime.fromisoformat(indexed_at).tzinfo is not None
    normalized = dict(result)
    normalized["indexed_at"] = "<indexed_at>"
    normalized_serialized = serialize_workbook_map(normalized)
    golden_path = Path(__file__).parents[1] / "golden" / f"{fixture_id.lower()}-workbook-map.json"
    golden_bytes = golden_path.read_bytes()
    budget_path = Path(__file__).parents[2] / "benchmarks" / "results" / "map-budgets.json"
    budget_document = cast(
        dict[str, object],
        json.loads(budget_path.read_text(encoding="utf-8")),
    )
    encoding_name = cast(str, budget_document["encoding"])
    recorded_tiktoken_version = cast(str, budget_document["tiktokenVersion"])
    encoding = tiktoken.get_encoding(encoding_name)
    budget = cast(
        dict[str, int | None],
        cast(dict[str, object], budget_document["maps"])[fixture_id],
    )

    assert golden_bytes == normalized_serialized.encode("utf-8") + b"\n"
    assert encoding_name == "o200k_base"
    assert distribution_version("tiktoken") == recorded_tiktoken_version
    assert budget["characterCap"] == 8_000
    assert len(normalized_serialized) == budget["characters"]
    assert len(encoding.encode(normalized_serialized)) == budget["tokens"]
    assert len(serialized) <= 8_000
    if fixture_id == "F03":
        assert budget["tokenCap"] == 1_500
        assert len(encoding.encode(serialized)) <= 1_500
        assert "GrowthRate" not in serialized
    else:
        assert budget["tokenCap"] is None
        assert result["sheets"] == 40
        assert result["namesMore"] == 280
        sheet_list = cast(list[dict[str, object]], result["sheetList"])
        assert sheet_list[-2]["vis"] == "hidden"
        assert sheet_list[-1]["vis"] == "veryHidden"


def test_every_current_fixture_map_obeys_the_character_cap(
    tmp_path: Path,
    p2_fixtures: dict[str, Path],
) -> None:
    for fixture_id, workbook in sorted(p2_fixtures.items()):
        result = build_workbook_map(workbook, index_dir=tmp_path / fixture_id)
        assert len(serialize_workbook_map(result)) <= 8_000, fixture_id


def _region_ranges(store: IndexStore, sheet_name: str) -> tuple[str, ...]:
    rows = store.connection.execute(
        """
        SELECT r.row_min, r.row_max, r.col_min, r.col_max
        FROM regions AS r JOIN sheets AS s ON s.id = r.sheet_id
        WHERE s.name = ? ORDER BY r.n
        """,
        (sheet_name,),
    ).fetchall()
    result: list[str] = []
    for row in rows:
        start = make_cell_ref(int(row["row_min"]), int(row["col_min"]))
        end = make_cell_ref(int(row["row_max"]), int(row["col_max"]))
        result.append(start if start == end else f"{start}:{end}")
    return tuple(result)
