"""Block-level dependency graph query and invariant tests."""

from __future__ import annotations

import sqlite3
from collections.abc import Iterator
from dataclasses import FrozenInstanceError
from pathlib import Path
from threading import Event, Thread
from typing import Any, cast

import pytest

import excel_lsp.core.graph.queries as graph_queries
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.graph import DependencyGraph, GraphArea, GraphTarget, TraceResult
from excel_lsp.core.index.edges import EdgeSchemaError, EdgeStore, SQLiteConnectionLike
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import Rect, SheetDescriptor


@pytest.mark.parametrize("prefer_rtree", [True, False])
def test_direct_queries_have_backend_parity_and_satisfy_i12_i14(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        graph = DependencyGraph(store.connection, store.edge_store)
        calc = GraphArea(2, "Calc", Rect(1, 2, 2, 2))

        precedents = graph.direct_precedents(calc)
        assert [(hop.target.ref, hop.via) for hop in precedents] == [("Inputs!A1:A1048576", "ref")]

        # I14: a whole-column edge matches a point at either Excel row bound.
        for row in (1, 1_048_576):
            dependents = graph.direct_dependents(GraphArea(1, "Inputs", Rect(row, row, 1, 1)))
            assert [hop.target.symbol for hop in dependents] == ["fblock:Calc:0"]

        # I12: following a block's precedents back through dependents recovers it.
        recovered = {
            dependent.target.symbol
            for precedent in precedents
            for dependent in graph.direct_dependents(precedent.target)
        }
        assert "fblock:Calc:0" in recovered


def test_packed_cell_sources_and_semantic_ordering(tmp_path: Path) -> None:
    with _graph_store(tmp_path) as store:
        _insert_edge(
            store,
            edge_id=4,
            src_kind="cell",
            src_id=(5 << 16) | 2,
            src_sheet_id=2,
            dst_sheet_id=1,
            rect=Rect(1, 1, 2, 2),
            via="name:Rate",
        )
        _insert_fblock(store, block_id=14, sheet_id=2, n=2, rect=Rect(4, 4, 2, 2))
        _insert_edge(
            store,
            edge_id=5,
            src_kind="fblock",
            src_id=14,
            src_sheet_id=2,
            dst_sheet_id=1,
            rect=Rect(1, 1, 2, 2),
            via="ref",
        )
        graph = DependencyGraph(store.connection, store.edge_store)

        dependents = graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 2, 2)))
        assert [hop.target.label for hop in dependents] == [
            "fblock:Calc:2",
            "cell:Calc!B5",
        ]
        assert (
            graph.direct_precedents(GraphArea(2, "Calc", Rect(5, 5, 2, 2)))[0].target.ref
            == "Inputs!B1"
        )


def test_trace_tree_is_breadth_first_counted_cycle_guarded_and_capped(tmp_path: Path) -> None:
    with _graph_store(tmp_path) as store:
        graph = DependencyGraph(store.connection, store.edge_store)
        source = GraphArea(1, "Inputs", Rect(1, 1, 1, 1))

        result = graph.trace_dependents(source, depth=3, max_nodes=20)
        assert result.node_count == 3
        assert result.edge_count == 2
        assert result.truncated is False
        assert result.root.child_count == 1
        calc = result.root.children[0]
        assert calc.target.symbol == "fblock:Calc:0"
        assert calc.child_count == 1
        assert calc.children[0].target.symbol == "fblock:Summary:0"

        capped = graph.trace_dependents(source, depth=8, max_nodes=2)
        assert (capped.node_count, capped.edge_count, capped.truncated) == (2, 1, True)
        assert capped.root.children[0].child_count == 1

        with pytest.raises(ValueError, match="depth"):
            graph.trace_dependents(source, depth=9)
        with pytest.raises(ValueError, match="max_nodes"):
            graph.trace_dependents(source, max_nodes=0)


def test_precedent_trace_keeps_ranges_and_opaque_sinks_without_expanding_cells(
    tmp_path: Path,
) -> None:
    with _graph_store(tmp_path) as store:
        _insert_edge(
            store,
            edge_id=4,
            src_kind="fblock",
            src_id=13,
            src_sheet_id=3,
            dst_sheet_id=None,
            rect=None,
            via="opaque:INDIRECT",
        )
        reads: list[str] = []

        def authorizer(
            action: int,
            arg1: str | None,
            _arg2: str | None,
            _database: str | None,
            _trigger: str | None,
        ) -> int:
            if action == sqlite3.SQLITE_READ and arg1 is not None:
                reads.append(arg1)
                if arg1 == "cells":
                    return sqlite3.SQLITE_DENY
            return sqlite3.SQLITE_OK

        store.connection.set_authorizer(authorizer)
        try:
            graph = DependencyGraph(store.connection, store.edge_store)
            direct_precedents = graph.direct_precedents(GraphArea(3, "Summary", Rect(1, 1, 3, 3)))
            direct_dependents = graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
            result = graph.trace_precedents(GraphArea(3, "Summary", Rect(1, 1, 3, 3)), depth=2)
            dependent_result = graph.trace_dependents(
                GraphArea(1, "Inputs", Rect(1, 1, 1, 1)), depth=2
            )
            path_result = graph.trace_path(
                GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
                GraphArea(3, "Summary", Rect(1, 1, 3, 3)),
            )
        finally:
            store.connection.set_authorizer(None)

        assert "cells" not in reads
        assert [hop.target.label for hop in direct_precedents] == [
            "Calc!B1:B2",
            "opaque:INDIRECT",
        ]
        assert direct_dependents[0].target.symbol == "fblock:Calc:0"
        assert [child.target.label for child in result.root.children] == [
            "Calc!B1:B2",
            "opaque:INDIRECT",
        ]
        assert result.root.children[1].children == ()
        assert result.root.children[0].children[0].target.ref == "Inputs!A1:A1048576"
        assert dependent_result.root.children[0].target.symbol == "fblock:Calc:0"
        assert path_result.connected is True


def test_trace_cycle_guard_keeps_a_cycle_as_a_finite_leaf(tmp_path: Path) -> None:
    with _graph_store(tmp_path) as store:
        _insert_edge(
            store,
            edge_id=4,
            src_kind="fblock",
            src_id=11,
            src_sheet_id=2,
            dst_sheet_id=2,
            rect=Rect(1, 2, 2, 2),
            via="ref",
        )
        result = DependencyGraph(store.connection, store.edge_store).trace_dependents(
            GraphArea(1, "Inputs", Rect(1, 1, 1, 1)), depth=8
        )

        calc = result.root.children[0]
        assert [child.target.symbol for child in calc.children] == [
            "fblock:Calc:0",
            "fblock:Summary:0",
        ]
        assert calc.children[0].children == ()
        assert result.node_count == 4


def test_trace_path_returns_bounded_shortest_paths_and_disconnected_result(
    tmp_path: Path,
) -> None:
    with _graph_store(tmp_path) as store:
        _insert_fblock(store, block_id=14, sheet_id=2, n=1, rect=Rect(3, 3, 2, 2))
        _insert_edge(
            store,
            edge_id=4,
            src_kind="fblock",
            src_id=14,
            src_sheet_id=2,
            dst_sheet_id=1,
            rect=Rect(1, 1, 1, 1),
            via="name:Input",
        )
        _insert_edge(
            store,
            edge_id=5,
            src_kind="fblock",
            src_id=13,
            src_sheet_id=3,
            dst_sheet_id=2,
            rect=Rect(3, 3, 2, 2),
            via="structured:Alternate",
        )
        graph = DependencyGraph(store.connection, store.edge_store)
        source = GraphArea(1, "Inputs", Rect(1, 1, 1, 1))
        destination = GraphArea(3, "Summary", Rect(1, 1, 3, 3))

        result = graph.trace_path(source, destination, max_paths=1)
        assert result.connected is True
        assert result.truncated is True
        assert len(result.paths) == 1
        assert [step.symbol for step in result.paths[0]] == [
            "cell:Inputs!A1",
            "fblock:Calc:0",
            "fblock:Summary:0",
        ]
        assert [step.via for step in result.paths[0]] == [None, "ref", "ref"]

        all_paths = graph.trace_path(source, destination, max_paths=3)
        assert len(all_paths.paths) == 2
        assert [path[1].symbol for path in all_paths.paths] == [
            "fblock:Calc:0",
            "fblock:Calc:1",
        ]
        assert graph.trace_path(source, destination, max_depth=1) == type(result)(False, ())
        assert graph.trace_path(GraphArea(1, "Inputs", Rect(9, 9, 9, 9)), destination) == type(
            result
        )(False, ())

        with pytest.raises(ValueError, match="max_paths"):
            graph.trace_path(source, destination, max_paths=0)
        with pytest.raises(ValueError, match="max_depth"):
            graph.trace_path(source, destination, max_depth=13)


def test_trace_path_enforces_global_exploration_cap(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    with _graph_store(tmp_path) as store:
        monkeypatch.setattr(graph_queries, "_MAX_PATH_NODES", 2)
        result = DependencyGraph(store.connection, store.edge_store).trace_path(
            GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
            GraphArea(3, "Summary", Rect(1, 1, 3, 3)),
        )

        assert result.connected is False
        assert result.paths == ()
        assert result.truncated is True


def test_trace_path_preserves_parallel_edges_with_distinct_reasons(tmp_path: Path) -> None:
    with _graph_store(tmp_path) as store:
        _insert_edge(
            store,
            edge_id=4,
            src_kind="fblock",
            src_id=11,
            src_sheet_id=2,
            dst_sheet_id=1,
            rect=Rect(1, 1, 1, 1),
            via="name:ParallelInput",
        )
        graph = DependencyGraph(store.connection, store.edge_store)
        result = graph.trace_path(
            GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
            GraphArea(3, "Summary", Rect(1, 1, 3, 3)),
            max_paths=3,
        )

        assert result.connected is True
        assert len(result.paths) == 2
        assert {tuple(step.via for step in path) for path in result.paths} == {
            (None, "ref", "ref"),
            (None, "name:ParallelInput", "ref"),
        }
        assert all(
            [step.symbol for step in path]
            == ["cell:Inputs!A1", "fblock:Calc:0", "fblock:Summary:0"]
            for path in result.paths
        )


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_small_trace_cap_has_near_constant_work_with_50k_irrelevant_edges(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        if prefer_rtree and store.edge_store.backend != "rtree":
            pytest.skip("SQLite RTree is unavailable")
        _insert_fblock(store, block_id=900, sheet_id=2, n=60_000, rect=Rect(500_000, 500_000, 2, 2))
        _insert_edge(
            store,
            edge_id=900,
            src_kind="fblock",
            src_id=900,
            src_sheet_id=2,
            dst_sheet_id=1,
            rect=Rect(2, 2, 2, 2),
            via="ref",
        )
        _insert_fblock(store, block_id=901, sheet_id=2, n=60_001, rect=Rect(500_001, 500_001, 2, 2))
        _insert_edge(
            store,
            edge_id=901,
            src_kind="fblock",
            src_id=901,
            src_sheet_id=2,
            dst_sheet_id=1,
            rect=Rect(2, 2, 2, 2),
            via="ref",
        )
        progress_by_fanout: dict[int, int] = {}
        previous = 0
        for fanout in (1_000, 10_000, 50_000):
            indexes = range(previous, fanout)
            with store.transaction():
                store.connection.executemany(
                    "INSERT INTO fblocks VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (
                        (
                            1_000 + index,
                            2,
                            10 + index,
                            "=RC",
                            10 + index,
                            10 + index,
                            2,
                            2,
                            0,
                            0,
                        )
                        for index in indexes
                    ),
                )
                store.connection.executemany(
                    """
                    INSERT INTO edges(
                        id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                        dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        (
                            1_000 + index,
                            "fblock",
                            1_000 + index,
                            2,
                            1,
                            3,
                            3,
                            3,
                            3,
                            "ref",
                        )
                        for index in range(previous, fanout)
                    ),
                )
                store.rebuild_graph_spatial_index()

            progress_calls = 0

            def progress() -> int:
                nonlocal progress_calls
                progress_calls += 1
                return 0

            store.connection.set_progress_handler(progress, 100)
            try:
                result = DependencyGraph(store.connection, store.edge_store).trace_dependents(
                    GraphArea(1, "Inputs", Rect(2, 2, 2, 2)), depth=1, max_nodes=2
                )
            finally:
                store.connection.set_progress_handler(None, 0)
            assert (result.node_count, result.edge_count, result.truncated) == (2, 1, True)
            progress_by_fanout[fanout] = progress_calls
            previous = fanout

        counts = tuple(progress_by_fanout.values())
        assert max(counts) <= min(counts) + 10, progress_by_fanout


@pytest.mark.parametrize("prefer_rtree", [True, False])
def test_ranked_spatial_lookup_finds_actual_semantic_prefix_without_temp_sort(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        area = GraphArea(1, "Inputs", Rect(1, 1, 1, 1))
        rank = store.edge_store.first_matching_rank("dependents", area.sheet_id, area.rect)
        assert rank == 1
        assert store.edge_store.edge_id_at_rank("dependents", area.sheet_id, area.rect, rank) == 1

        if store.edge_store.backend == "rtree":
            sql = """
                EXPLAIN QUERY PLAN SELECT 1 FROM edge_rtree
                WHERE sheet_min <= 1 AND sheet_max >= 1
                  AND row_min <= 1 AND row_max >= 1
                  AND col_min <= 1 AND col_max >= 1
                  AND rank_min >= 1 AND rank_max <= 1 LIMIT 1
            """
        else:
            sql = """
                EXPLAIN QUERY PLAN SELECT MIN(rank) FROM edge_intervals
                WHERE sheet_id = 1 AND rank > 0
                  AND row_min <= 1 AND row_max >= 1
                  AND col_min <= 1 AND col_max >= 1
            """
        plan = " ".join(str(row[3]) for row in store.connection.execute(sql))
        assert "TEMP B-TREE" not in plan
        assert (
            "VIRTUAL TABLE INDEX" in plan
            if store.edge_store.backend == "rtree"
            else "edge_intervals_overlap" in plan
        )


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_small_precedent_cap_has_near_constant_work_with_50k_irrelevant_edges(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        if prefer_rtree and store.edge_store.backend != "rtree":
            pytest.skip("SQLite RTree is unavailable")
        _insert_edge(
            store,
            edge_id=900,
            src_kind="fblock",
            src_id=11,
            src_sheet_id=2,
            dst_sheet_id=1,
            rect=Rect(500_000, 500_000, 3, 3),
            via="ref",
        )
        progress_by_fanout: dict[int, int] = {}
        previous = 0
        for fanout in (1_000, 10_000, 50_000):
            with store.transaction():
                store.connection.executemany(
                    "INSERT INTO fblocks VALUES (?, 2, ?, '=RC', ?, ?, 2, 2, 0, 0)",
                    (
                        (100_000 + index, 100_000 + index, 10 + index, 10 + index)
                        for index in range(previous, fanout)
                    ),
                )
                store.connection.executemany(
                    """
                    INSERT INTO edges(
                        id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                        dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        (
                            100_000 + index,
                            "fblock",
                            100_000 + index,
                            2,
                            1,
                            10 + index,
                            10 + index,
                            3,
                            3,
                            "ref",
                        )
                        for index in range(previous, fanout)
                    ),
                )
                store.rebuild_graph_spatial_index()

            progress_calls = 0

            def progress() -> int:
                nonlocal progress_calls
                progress_calls += 1
                return 0

            store.connection.set_progress_handler(progress, 100)
            try:
                result = DependencyGraph(store.connection, store.edge_store).trace_precedents(
                    GraphArea(2, "Calc", Rect(1, 2, 2, 2)), depth=1, max_nodes=2
                )
            finally:
                store.connection.set_progress_handler(None, 0)
            assert (result.node_count, result.edge_count, result.truncated) == (2, 1, True)
            progress_by_fanout[fanout] = progress_calls
            previous = fanout

        counts = tuple(progress_by_fanout.values())
        assert max(counts) <= min(counts) + 10, progress_by_fanout


@pytest.mark.parametrize("prefer_rtree", [True, False])
def test_fractional_relational_and_spatial_coordinates_are_corrupt(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (
                99, 'cell', 65537, 1, 1, 1.5, 2, 1, 1, 'ref'
            )
            """
        )
        with pytest.raises(ExcelLSPError, match="not an integer") as relational:
            store.rebuild_graph_spatial_index()
        assert relational.value.code is ErrorCode.CORRUPT

    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (
                99, 'cell', 65537, 1, 1, 1, 1, 1, 1, 'ref'
            )
            """
        )
        if store.edge_store.backend == "rtree":
            store.connection.execute("INSERT INTO edge_rtree VALUES (99, 1, 1, 1, 1.5, 1, 1, 1, 1)")
        else:
            store.connection.execute("INSERT INTO edge_intervals VALUES (99, 1, 1, 1.5, 1, 1, 1)")
            with pytest.raises(ValueError, match="not an integer"):
                store.edge_store.ranked_mirror(99, "dependents")
        with pytest.raises(ExcelLSPError, match="spatial mirrors") as spatial:
            DependencyGraph(store.connection, store.edge_store).direct_precedents(
                GraphArea(1, "Inputs", Rect(1, 1, 1, 1))
            )
        assert spatial.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False])
def test_negative_formula_block_ordinals_are_corrupt_not_raw_value_errors(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        _insert_fblock(store, block_id=99, sheet_id=2, n=-1, rect=Rect(10, 10, 2, 2))
        with pytest.raises(ExcelLSPError, match="ordinal is negative") as captured:
            store.connection.execute(
                """
                INSERT INTO edges(
                    id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES (99, 'fblock', 99, 2, 1, 1, 1, 1, 1, 'ref')
                """
            )
            store.rebuild_graph_spatial_index()
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False])
def test_touched_edges_require_bidirectional_spatial_mirror_integrity(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        graph = DependencyGraph(store.connection, store.edge_store)
        store.edge_store.delete(1)
        with pytest.raises(ExcelLSPError, match="spatial mirror") as missing:
            graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
        assert missing.value.code is ErrorCode.CORRUPT

    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        graph = DependencyGraph(store.connection, store.edge_store)
        store.edge_store.insert(99, 1, Rect(1, 1, 1, 1))
        with pytest.raises(ExcelLSPError, match="spatial mirror") as orphan:
            graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
        assert orphan.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    ("direction", "corruption"),
    (
        ("source", "rectangle"),
        ("source", "rank"),
        ("source", "missing"),
        ("destination", "rectangle"),
        ("destination", "rank"),
        ("destination", "missing"),
    ),
)
def test_clean_but_corrupt_returned_edge_mirrors_are_rejected(
    tmp_path: Path,
    prefer_rtree: bool,
    direction: str,
    corruption: str,
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        table = (
            store.edge_store.source_table_name
            if direction == "source"
            else store.edge_store.table_name
        )
        if corruption == "missing":
            store.connection.execute(f"DELETE FROM {table} WHERE edge_id = 1")
        elif corruption == "rank":
            if store.edge_store.backend == "rtree":
                store.connection.execute(
                    f"UPDATE {table} SET rank_min = 99, rank_max = 99 WHERE edge_id = 1"
                )
            else:
                store.connection.execute(f"UPDATE {table} SET rank = 99 WHERE edge_id = 1")
        else:
            store.connection.execute(f"UPDATE {table} SET row_min = 2 WHERE edge_id = 1")
        # Simulate an externally tampered sidecar whose persisted clean marker
        # and integrity epoch have both been forged. The per-edge validator
        # remains a separate defense after the O(1) trust gate.
        store.connection.execute(
            """
            UPDATE graph_spatial_state
            SET dirty = 0, clean_epoch = mutation_epoch
            WHERE singleton = 1
            """
        )

        with pytest.raises(ExcelLSPError, match=r"spatial mirror") as captured:
            if direction == "source":
                # The intact destination mirror returns edge 1; validation must
                # independently reject its damaged source mirror.
                store.dependency_graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
            else:
                # Relational source overlap returns edge 1 without consulting
                # the damaged destination mirror first.
                store.dependency_graph.direct_precedents(GraphArea(2, "Calc", Rect(1, 1, 2, 2)))
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("direction", ["source", "destination"])
def test_clean_orphan_mirrors_are_rejected_by_relational_cross_check(
    tmp_path: Path, prefer_rtree: bool, direction: str
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        state_column = "precedent_rank_max" if direction == "source" else "dependent_rank_max"
        # Reuse a dense rank whose legitimate mirror is on another sheet: rank
        # 1 for the source mirror and rank 2 for the destination mirror. This
        # keeps open-time rank validation sound while reaching the deeper
        # orphan-row validator.
        rank_max = int(
            store.connection.execute(
                f"SELECT {state_column} FROM graph_spatial_state WHERE singleton = 1"
            ).fetchone()[0]
        )
        assert rank_max == 2
        rank = 1 if direction == "source" else 2
        table = (
            store.edge_store.source_table_name
            if direction == "source"
            else store.edge_store.table_name
        )
        sheet_id = 2 if direction == "source" else 1
        rect = Rect(1, 2, 2, 2) if direction == "source" else Rect(1, 1_048_576, 1, 1)
        if store.edge_store.backend == "rtree":
            store.connection.execute(
                f"""
                INSERT INTO {table}(
                    edge_id, sheet_min, sheet_max, row_min, row_max,
                    col_min, col_max, rank_min, rank_max
                ) VALUES (99, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    sheet_id,
                    sheet_id,
                    rect.row_min,
                    rect.row_max,
                    rect.col_min,
                    rect.col_max,
                    rank,
                    rank,
                ),
            )
        else:
            store.connection.execute(
                f"""
                INSERT INTO {table}(
                    edge_id, sheet_id, row_min, row_max, col_min, col_max, rank
                ) VALUES (99, ?, ?, ?, ?, ?, ?)
                """,
                (
                    sheet_id,
                    rect.row_min,
                    rect.row_max,
                    rect.col_min,
                    rect.col_max,
                    rank,
                ),
            )
        store.connection.execute(
            """
            UPDATE graph_spatial_state
            SET dirty = 0, clean_epoch = mutation_epoch
            WHERE singleton = 1
            """
        )

        # Persisted facade opening now validates the complete mirror projection,
        # so an orphan is rejected before any bounded query can consult it.
        with pytest.raises(EdgeSchemaError, match=r"persisted .* spatial mirror"):
            EdgeStore(store.connection)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    ("surface", "direction", "edge_id"),
    (
        ("trace_precedents", "precedents", 2),
        ("trace_dependents", "dependents", 1),
        ("trace_path", "dependents", 1),
    ),
)
def test_clean_marker_cannot_hide_missing_queried_direction_from_bounded_surfaces(
    tmp_path: Path,
    prefer_rtree: bool,
    surface: str,
    direction: str,
    edge_id: int,
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        table = (
            store.edge_store.source_table_name
            if direction == "precedents"
            else store.edge_store.table_name
        )
        store.connection.execute(f"DELETE FROM {table} WHERE edge_id = ?", (edge_id,))
        mutation_epoch = int(
            store.connection.execute(
                "SELECT mutation_epoch FROM graph_spatial_state WHERE singleton = 1"
            ).fetchone()[0]
        )
        store.connection.execute("UPDATE graph_spatial_state SET dirty = 0 WHERE singleton = 1")
        clean_epoch = int(
            store.connection.execute(
                "SELECT clean_epoch FROM graph_spatial_state WHERE singleton = 1"
            ).fetchone()[0]
        )
        assert mutation_epoch > clean_epoch

        # Direct construction shapes the full persisted-mirror rejection as a
        # public corruption error rather than leaking EdgeSchemaError.
        with pytest.raises(ExcelLSPError) as captured:
            graph = DependencyGraph(store.connection)
            if surface == "trace_precedents":
                graph.trace_precedents(GraphArea(3, "Summary", Rect(1, 1, 3, 3)))
            elif surface == "trace_dependents":
                graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
            else:
                graph.trace_path(
                    GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
                    GraphArea(3, "Summary", Rect(1, 1, 3, 3)),
                )
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("timing", ["pre-open", "post-open"])
@pytest.mark.parametrize("corruption", ["missing", "displaced"])
@pytest.mark.parametrize(
    ("surface", "direction", "edge_id"),
    (
        ("trace_precedents", "precedents", 2),
        ("trace_dependents", "dependents", 1),
        ("trace_path", "dependents", 1),
    ),
)
def test_forged_clean_epoch_cannot_hide_queried_mirror_tampering(
    tmp_path: Path,
    prefer_rtree: bool,
    timing: str,
    corruption: str,
    surface: str,
    direction: str,
    edge_id: int,
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        graph = (
            DependencyGraph(store.connection, store.edge_store) if timing == "post-open" else None
        )
        table = (
            store.edge_store.source_table_name
            if direction == "precedents"
            else store.edge_store.table_name
        )
        if corruption == "missing":
            store.connection.execute(f"DELETE FROM {table} WHERE edge_id = ?", (edge_id,))
        else:
            store.connection.execute(
                f"""
                UPDATE {table}
                SET col_min = col_min + 1, col_max = col_max + 1
                WHERE edge_id = ?
                """,
                (edge_id,),
            )
        store.connection.execute(
            """
            UPDATE graph_spatial_state
            SET dirty = 0, clean_epoch = mutation_epoch
            WHERE singleton = 1
            """
        )
        forged = store.connection.execute(
            """
            SELECT dirty, mutation_epoch, clean_epoch
            FROM graph_spatial_state WHERE singleton = 1
            """
        ).fetchone()
        assert forged is not None
        assert (int(forged[0]), int(forged[1]) == int(forged[2])) == (0, True)

        with pytest.raises(ExcelLSPError) as captured:
            active_graph = graph or DependencyGraph(store.connection)
            if surface == "trace_precedents":
                active_graph.trace_precedents(GraphArea(3, "Summary", Rect(1, 1, 3, 3)))
            elif surface == "trace_dependents":
                active_graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
            else:
                active_graph.trace_path(
                    GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
                    GraphArea(3, "Summary", Rect(1, 1, 3, 3)),
                )
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize(
    ("shadow_table", "surface"),
    (
        ("edge_rtree_node", "direct_dependents"),
        ("edge_rtree_node", "trace_dependents"),
        ("edge_rtree_node", "trace_path"),
        ("edge_source_rtree_rowid", "direct_precedents"),
        ("edge_source_rtree_rowid", "trace_precedents"),
    ),
)
def test_post_open_rtree_shadow_corruption_is_shaped_at_every_public_query_boundary(
    tmp_path: Path, shadow_table: str, surface: str
) -> None:
    with _graph_store(tmp_path, prefer_rtree=True) as store:
        if store.edge_store.backend != "rtree":
            pytest.skip("SQLite RTree is unavailable")
        graph = DependencyGraph(store.connection, store.edge_store)
        sealed_state = tuple(
            int(value)
            for value in store.connection.execute(
                """
                SELECT dirty, mutation_epoch, clean_epoch
                FROM graph_spatial_state WHERE singleton = 1
                """
            ).fetchone()
        )
        assert sealed_state[0] == 0
        assert sealed_state[1] == sealed_state[2]
        store.edge_store.require_clean()

        store.connection.execute(f"DELETE FROM {shadow_table}")
        store.connection.execute(
            """
            UPDATE graph_spatial_state
            SET dirty = ?, mutation_epoch = ?, clean_epoch = ?
            WHERE singleton = 1
            """,
            sealed_state,
        )
        restored_state = tuple(
            int(value)
            for value in store.connection.execute(
                """
                SELECT dirty, mutation_epoch, clean_epoch
                FROM graph_spatial_state WHERE singleton = 1
                """
            ).fetchone()
        )
        assert restored_state == sealed_state

        with pytest.raises(ExcelLSPError) as captured:
            _invoke_graph_surface(graph, surface)
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_live_relational_iteration_database_errors_are_shaped_at_public_boundary(
    tmp_path: Path, surface: str
) -> None:
    with _graph_store(tmp_path) as store:
        graph = DependencyGraph(store.connection, store.edge_store)

        def deny_sheet_reads(
            action: int,
            arg1: str | None,
            _arg2: str | None,
            _database: str | None,
            _trigger: str | None,
        ) -> int:
            if action == sqlite3.SQLITE_READ and arg1 == "sheets":
                return sqlite3.SQLITE_DENY
            return sqlite3.SQLITE_OK

        store.connection.set_authorizer(deny_sheet_reads)
        try:
            with pytest.raises(ExcelLSPError) as captured:
                _invoke_graph_surface(graph, surface)
        finally:
            store.connection.set_authorizer(None)
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_closed_connection_programming_errors_escape_every_public_query_boundary(
    tmp_path: Path, surface: str
) -> None:
    store = _graph_store(tmp_path)
    connection = store.connection
    graph = DependencyGraph(connection, store.edge_store)
    store.close()

    with pytest.raises(sqlite3.ProgrammingError) as reference:
        connection.execute("SELECT 1")
    with pytest.raises(sqlite3.ProgrammingError) as captured:
        _invoke_graph_surface(graph, surface)

    assert type(captured.value) is sqlite3.ProgrammingError
    assert str(captured.value) == str(reference.value)


def test_closed_connection_programming_error_escapes_graph_construction() -> None:
    connection = sqlite3.connect(":memory:")
    connection.close()

    with pytest.raises(sqlite3.ProgrammingError) as reference:
        connection.execute("SELECT 1")
    with pytest.raises(sqlite3.ProgrammingError) as captured:
        DependencyGraph(connection)

    assert type(captured.value) is sqlite3.ProgrammingError
    assert str(captured.value) == str(reference.value)


@pytest.mark.parametrize(
    "marker",
    (
        EdgeSchemaError("injected edge schema failure"),
        sqlite3.OperationalError("injected SQLite storage failure"),
    ),
)
def test_graph_construction_still_shapes_storage_failures_as_corrupt(
    monkeypatch: pytest.MonkeyPatch, marker: BaseException
) -> None:
    connection = sqlite3.connect(":memory:")

    def fail_edge_store(_connection: sqlite3.Connection) -> None:
        raise marker

    monkeypatch.setattr(graph_queries, "EdgeStore", fail_edge_store)
    try:
        with pytest.raises(ExcelLSPError) as captured:
            DependencyGraph(connection)
    finally:
        connection.close()

    assert captured.value.code is ErrorCode.CORRUPT
    assert captured.value.__cause__ is marker


@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
@pytest.mark.parametrize("error_kind", ["type-error", "excel-lsp-error"])
def test_public_query_boundary_preserves_non_database_errors(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    surface: str,
    error_kind: str,
) -> None:
    with _graph_store(tmp_path) as store:
        graph = DependencyGraph(store.connection, store.edge_store)
        if error_kind == "type-error":
            marker = TypeError(f"programming fault from {surface}")

            def fail() -> None:
                raise marker

            monkeypatch.setattr(graph, "_require_clean_spatial", fail)
            with pytest.raises(TypeError) as captured:
                _invoke_graph_surface(graph, surface)
        else:
            marker = ExcelLSPError(ErrorCode.CORRUPT, f"existing error from {surface}")

            def fail() -> None:
                raise marker

            monkeypatch.setattr(graph, "_require_clean_spatial", fail)
            with pytest.raises(ExcelLSPError) as captured:
                _invoke_graph_surface(graph, surface)

        assert captured.value is marker


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_public_graph_operation_keeps_one_snapshot_across_live_writer_commit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    surface: str,
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as reader:
        graph = reader.dependency_graph
        expected = _invoke_graph_surface(graph, surface)
        with IndexStore(reader.path, prefer_rtree=prefer_rtree) as writer:
            original_require_clean = graph._require_clean_spatial
            mutation_committed = False

            def validate_then_commit_mutation() -> None:
                nonlocal mutation_committed
                original_require_clean()
                if mutation_committed:
                    return
                mutation_committed = True
                edge_id = 2 if surface == "trace_precedents" else 1
                with writer.transaction():
                    writer.connection.execute("DELETE FROM edges WHERE id = ?", (edge_id,))
                    writer.rebuild_graph_spatial_index()

            monkeypatch.setattr(graph, "_require_clean_spatial", validate_then_commit_mutation)
            assert _invoke_graph_surface(graph, surface) == expected
            assert mutation_committed is True
            assert reader.connection.in_transaction is False

            # The graph-owned snapshot is gone. Reuse sees the writer's newer
            # trust tuple and shapes it as corruption instead of mixing reads.
            with pytest.raises(ExcelLSPError) as captured:
                _invoke_graph_surface(graph, surface)
            assert captured.value.code is ErrorCode.CORRUPT
            assert reader.connection.in_transaction is False


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_delete_journal_writer_commit_waits_for_public_graph_snapshot(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    surface: str,
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as reader:
        assert reader.connection.execute("PRAGMA journal_mode = DELETE").fetchone()[0] == "delete"

        writer_ready = Event()
        start_writer = Event()
        writes_ready = Event()
        commit_attempted = Event()
        commit_returned = Event()
        writer_finished = Event()
        writer_errors: list[BaseException] = []

        def mutate_graph() -> None:
            try:
                with IndexStore(reader.path, prefer_rtree=prefer_rtree) as writer:
                    assert (
                        writer.connection.execute("PRAGMA journal_mode = DELETE").fetchone()[0]
                        == "delete"
                    )
                    writer._connection = _CommitObserver(  # type: ignore[assignment]
                        writer._connection,
                        commit_attempted,
                        commit_returned,
                    )
                    writer_ready.set()
                    if not start_writer.wait(10):
                        raise TimeoutError("reader never released the DELETE-mode writer")
                    edge_id = 2 if surface == "trace_precedents" else 1
                    with writer.transaction():
                        writer.connection.execute("DELETE FROM edges WHERE id = ?", (edge_id,))
                        writer.rebuild_graph_spatial_index()
                        writes_ready.set()
            except BaseException as exc:
                writer_errors.append(exc)
            finally:
                writer_ready.set()
                writes_ready.set()
                writer_finished.set()

        writer_thread = Thread(target=mutate_graph, name=f"graph-writer-{surface}")
        writer_thread.start()
        assert writer_ready.wait(10), "DELETE-mode writer did not initialize"
        assert writer_errors == []
        # Journal-mode setup on either SQLite handle can advance the reader's
        # data-version token. Seal only after both pagers are final, before the
        # writer is allowed to mutate graph storage.
        graph = DependencyGraph(reader.connection)
        expected = _invoke_graph_surface(graph, surface)
        original_require_clean = graph._require_clean_spatial
        writer_started = False

        def validate_then_start_writer() -> None:
            nonlocal writer_started
            original_require_clean()
            if writer_started:
                return
            writer_started = True
            start_writer.set()
            assert writes_ready.wait(10), "writer did not rebuild inside its transaction"
            assert commit_attempted.wait(10), "writer never attempted its DELETE-mode commit"
            assert commit_returned.wait(0.1) is False
            assert writer_finished.is_set() is False

        try:
            monkeypatch.setattr(graph, "_require_clean_spatial", validate_then_start_writer)
            assert _invoke_graph_surface(graph, surface) == expected
            assert reader.connection.in_transaction is False
            assert writer_finished.wait(10), "writer stayed blocked after the reader snapshot ended"
            assert commit_returned.is_set() is True
        finally:
            start_writer.set()
            writer_thread.join(10)

        assert writer_thread.is_alive() is False
        assert writer_errors == []
        assert _invoke_graph_surface(DependencyGraph(reader.connection), surface) != expected


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("begin_timing", ["before", "after"])
@pytest.mark.parametrize("release_mode", ["normal", "deny"])
@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_public_graph_begin_failure_is_released_or_conclusively_closed(
    tmp_path: Path,
    prefer_rtree: bool,
    journal_mode: str,
    begin_timing: str,
    release_mode: str,
    surface: str,
) -> None:
    seed = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    database = seed.path
    expected = _invoke_graph_surface(seed.dependency_graph, surface)
    seed.close()

    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_BeginFailureConnection,
    )
    assert isinstance(connection, _BeginFailureConnection)
    assert str(connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()[0]) == (
        journal_mode
    )
    graph = DependencyGraph(connection, EdgeStore(connection))
    connection.begin_timing = begin_timing
    connection.begin_marker = sqlite3.DatabaseError(f"BEGIN {begin_timing} effect for {surface}")

    if release_mode == "deny":

        def deny_rollback(
            action: int,
            argument_one: str | None,
            argument_two: str | None,
            database_name: str | None,
            trigger: str | None,
        ) -> int:
            del argument_two, database_name, trigger
            if action == sqlite3.SQLITE_TRANSACTION and argument_one == "ROLLBACK":
                return sqlite3.SQLITE_DENY
            return sqlite3.SQLITE_OK

        connection.set_authorizer(deny_rollback)

    with pytest.raises(ExcelLSPError) as captured:
        _invoke_graph_surface(graph, surface)
    assert captured.value.code is ErrorCode.CORRUPT
    assert captured.value.__cause__ is connection.begin_marker

    closed = begin_timing == "after" and release_mode == "deny"
    if closed:
        with pytest.raises(sqlite3.ProgrammingError):
            _ = connection.in_transaction
        with pytest.raises(sqlite3.ProgrammingError, match="poisoned"):
            _invoke_graph_surface(graph, surface)
    else:
        assert connection.in_transaction is False
        connection.set_authorizer(None)
        assert _invoke_graph_surface(graph, surface) == expected

    writer = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.execute(
            "INSERT INTO meta(key, value) VALUES ('begin_failure_writer', ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (surface,),
        )
        writer.commit()
    finally:
        writer.close()
        sqlite3.Connection.close(connection)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("state_spoof", ["false-positive", "false-negative"])
@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_poison_close_uses_native_state_when_subclass_spoofs_in_transaction(
    tmp_path: Path,
    prefer_rtree: bool,
    state_spoof: str,
    surface: str,
) -> None:
    seed = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    database = seed.path
    seed.close()
    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_SpoofedNativeStateConnection,
    )
    assert isinstance(connection, _SpoofedNativeStateConnection)
    assert connection.execute("PRAGMA journal_mode = DELETE").fetchone()[0] == "delete"
    graph = DependencyGraph(connection, EdgeStore(connection))
    connection.spoof_after_begin = state_spoof
    connection.close_timing = "before" if state_spoof == "false-positive" else "after"
    connection.set_authorizer(_deny_graph_rollback)

    with pytest.raises(ExcelLSPError) as captured:
        _invoke_graph_surface(graph, surface)
    assert captured.value.code is ErrorCode.CORRUPT
    assert connection.close_calls == (3 if state_spoof == "false-positive" else 1)
    with pytest.raises(sqlite3.ProgrammingError):
        sqlite3.Connection.execute(connection, "SELECT 1")
    if state_spoof == "false-negative":
        with pytest.raises(RuntimeError, match="virtual state probe"):
            _ = connection.in_transaction
    with pytest.raises(sqlite3.ProgrammingError, match="poisoned"):
        _invoke_graph_surface(graph, surface)

    writer = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.rollback()
    finally:
        writer.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("body_chain", ["cause", "context"])
def test_snapshot_cleanup_does_not_duplicate_prior_already_used_as_cleanup_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    body_chain: str,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    graph = store.dependency_graph
    marker = ValueError("graph body primary")
    prior_error = RuntimeError(f"shared graph {body_chain} evidence")

    def fail_query_body() -> None:
        if body_chain == "cause":
            raise marker from prior_error
        try:
            raise prior_error
        except RuntimeError:
            raise marker  # noqa: B904 - exercise implicit exception context

    def reuse_prior_as_rollback_error() -> None:
        raise prior_error

    with monkeypatch.context() as patch:
        patch.setattr(graph, "_require_clean_spatial", fail_query_body)
        patch.setattr(graph, "_rollback_owned_read_snapshot", reuse_prior_as_rollback_error)
        store.connection.set_authorizer(_deny_graph_rollback)
        with pytest.raises(ValueError) as captured:
            _invoke_graph_surface(graph, "direct_precedents")

    assert captured.value is marker
    assert isinstance(captured.value.__cause__, BaseExceptionGroup)
    assert captured.value.__cause__.exceptions[0] is prior_error
    _assert_unique_exception_identities(captured.value)
    with pytest.raises(sqlite3.ProgrammingError):
        sqlite3.Connection.execute(store._connection, "SELECT 1")
    store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("body_chain", ["cause", "context"])
def test_snapshot_cleanup_normalizes_recursive_group_membership_aliases(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    body_chain: str,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    graph = store.dependency_graph
    marker = ValueError("nested graph body primary")
    prior_error = RuntimeError(f"nested graph {body_chain} prior")
    cleanup_group, inner_group, distinct_one, distinct_two = _nested_cleanup_group(prior_error)

    def fail_query_body() -> None:
        if body_chain == "cause":
            raise marker from prior_error
        try:
            raise prior_error
        except RuntimeError:
            raise marker  # noqa: B904 - exercise implicit exception context

    def fail_rollback_with_nested_group() -> None:
        raise cleanup_group

    with monkeypatch.context() as patch:
        patch.setattr(graph, "_require_clean_spatial", fail_query_body)
        patch.setattr(graph, "_rollback_owned_read_snapshot", fail_rollback_with_nested_group)
        with pytest.raises(ValueError) as captured:
            _invoke_graph_surface(graph, "direct_precedents")

    assert captured.value is marker
    assert captured.value.__cause__ is cleanup_group
    assert cleanup_group.exceptions == (inner_group, distinct_two)
    assert inner_group.exceptions == (prior_error, distinct_one)
    assert distinct_two.__context__ is None
    _assert_unique_exception_identities(captured.value)
    assert store.connection.in_transaction is False
    store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_successful_snapshot_cleanup_detaches_nested_links_to_primary_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    surface: str,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    graph = store.dependency_graph
    primary_error = RuntimeError("first successful-snapshot cleanup failure")
    linked_error = RuntimeError("later cleanup member")
    linked_error.__context__ = primary_error
    linked_error.__suppress_context__ = True
    distinct_error = RuntimeError("distinct later cleanup evidence")
    inner_group = BaseExceptionGroup("nested later cleanup", (linked_error, distinct_error))
    cleanup_group = BaseExceptionGroup("outer later cleanup", (inner_group,))
    original_release = graph._release_owned_read_snapshot

    def release_then_report_failures() -> tuple[BaseException, ...]:
        assert original_release() == ()
        return (primary_error, cleanup_group)

    with monkeypatch.context() as patch:
        patch.setattr(graph, "_release_owned_read_snapshot", release_then_report_failures)
        with pytest.raises(RuntimeError) as captured:
            _invoke_graph_surface(graph, surface)

    assert captured.value is primary_error
    assert captured.value.__cause__ is cleanup_group
    assert inner_group.exceptions == (linked_error, distinct_error)
    assert linked_error.__context__ is None
    _assert_unique_exception_identities(captured.value)
    assert store.connection.in_transaction is False
    store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("surface", ["trace_precedents", "trace_dependents", "trace_path"])
def test_split_duplicate_rank_is_corrupt_after_exact_seal_restoration(
    tmp_path: Path,
    prefer_rtree: bool,
    surface: str,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    _insert_edge(
        store,
        edge_id=3,
        src_kind="fblock",
        src_id=11,
        src_sheet_id=2,
        dst_sheet_id=1,
        rect=Rect(1, 1_048_576, 1, 1),
        via="ref",
    )
    graph = store.dependency_graph
    valid_trace = _invoke_graph_surface(graph, surface)
    assert valid_trace is not None
    valid_direct = graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
    assert [(hop.target.symbol, hop.via) for hop in valid_direct] == [("fblock:Calc:0", "ref")]
    sealed_state = tuple(
        store.connection.execute(
            """
            SELECT singleton, dirty, dependent_rank_max, precedent_rank_max,
                   revision, mutation_epoch, clean_epoch
            FROM graph_spatial_state
            """
        ).fetchone()
    )
    sealed_catalog = tuple(
        tuple(row)
        for row in store.connection.execute(
            "SELECT direction, rank, key_text FROM graph_rank_keys ORDER BY direction, rank"
        )
    )

    store.connection.execute("UPDATE edges SET via = 'split-ref' WHERE id = 3")
    store.connection.executemany(
        "INSERT OR REPLACE INTO graph_rank_keys(direction, rank, key_text) VALUES (?, ?, ?)",
        sealed_catalog,
    )
    store.connection.execute(
        """
        UPDATE graph_spatial_state
        SET singleton = ?, dirty = ?, dependent_rank_max = ?, precedent_rank_max = ?,
            revision = ?, mutation_epoch = ?, clean_epoch = ?
        """,
        sealed_state,
    )

    with pytest.raises(ExcelLSPError) as direct_corrupt:
        graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
    assert direct_corrupt.value.code is ErrorCode.CORRUPT
    with pytest.raises(ExcelLSPError) as captured:
        _invoke_graph_surface(graph, surface)
    assert captured.value.code is ErrorCode.CORRUPT
    assert "stale or corrupt" in captured.value.message
    store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_live_graph_rejects_cross_connection_full_rank_and_seal_restoration(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    _insert_edge(
        store,
        edge_id=3,
        src_kind="fblock",
        src_id=11,
        src_sheet_id=2,
        dst_sheet_id=1,
        rect=Rect(1, 1_048_576, 1, 1),
        via="ref",
    )
    graph = store.dependency_graph
    assert _invoke_graph_surface(graph, "trace_dependents") is not None
    sealed_state = tuple(
        store.connection.execute(
            """
            SELECT singleton, dirty, dependent_rank_max, precedent_rank_max,
                   revision, mutation_epoch, clean_epoch
            FROM graph_spatial_state
            """
        ).fetchone()
    )
    sealed_catalog = tuple(
        tuple(row)
        for row in store.connection.execute(
            "SELECT direction, rank, key_text FROM graph_rank_keys ORDER BY direction, rank"
        )
    )

    writer = sqlite3.connect(store.path, timeout=5.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.execute("UPDATE edges SET via = 'cross-split' WHERE id = 3")
        writer.executemany(
            "INSERT OR REPLACE INTO graph_rank_keys(direction, rank, key_text) VALUES (?, ?, ?)",
            sealed_catalog,
        )
        writer.execute(
            """
            UPDATE graph_spatial_state
            SET singleton = ?, dirty = ?, dependent_rank_max = ?, precedent_rank_max = ?,
                revision = ?, mutation_epoch = ?, clean_epoch = ?
            """,
            sealed_state,
        )
        writer.commit()
    finally:
        writer.close()

    for surface in (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ):
        with pytest.raises(ExcelLSPError) as captured:
            _invoke_graph_surface(graph, surface)
        assert captured.value.code is ErrorCode.CORRUPT
    with pytest.raises(EdgeSchemaError, match="canonical"):
        EdgeStore(store.connection)
    store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_live_graph_rejects_trigger_removal_before_silent_relational_mutation(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    _insert_edge(
        store,
        edge_id=3,
        src_kind="fblock",
        src_id=11,
        src_sheet_id=2,
        dst_sheet_id=1,
        rect=Rect(1, 1_048_576, 1, 1),
        via="ref",
    )
    graph = store.dependency_graph
    assert _invoke_graph_surface(graph, "trace_dependents") is not None

    store.connection.execute("DROP TRIGGER edges_graph_spatial_dirty_update")
    store.connection.execute("UPDATE edges SET via = 'silent-split' WHERE id = 3")

    for surface in (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ):
        with pytest.raises(ExcelLSPError) as captured:
            _invoke_graph_surface(graph, surface)
        assert captured.value.code is ErrorCode.CORRUPT
    with pytest.raises(EdgeSchemaError, match=r"canonical|trigger"):
        EdgeStore(store.connection)
    store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("owner", ["raw", "store"], ids=["caller", "managed"])
@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_public_graph_operation_never_releases_caller_transaction(
    tmp_path: Path, prefer_rtree: bool, owner: str, surface: str
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        graph = store.dependency_graph
        expected = _invoke_graph_surface(graph, surface)
        if owner == "raw":
            store.connection.execute("BEGIN")
            try:
                assert _invoke_graph_surface(graph, surface) == expected
                assert store.connection.in_transaction is True
            finally:
                store.connection.rollback()
        else:
            with store.transaction():
                assert _invoke_graph_surface(graph, surface) == expected
                assert store.connection.in_transaction is True
                store.set_meta("snapshot_owner", surface)
            assert store.get_meta("snapshot_owner") == surface

        assert store.connection.in_transaction is False
        assert _invoke_graph_surface(graph, surface) == expected


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("owner", ["raw", "store"], ids=["caller", "managed"])
@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_public_graph_errors_preserve_caller_transaction_and_reuse(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    owner: str,
    surface: str,
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        graph = store.dependency_graph
        expected = _invoke_graph_surface(graph, surface)
        original_iter_rows = graph._iter_rows

        def exercise_errors() -> None:
            markers: tuple[BaseException, ...] = (
                sqlite3.DatabaseError(f"database failure from {surface}"),
                sqlite3.ProgrammingError(f"programming failure from {surface}"),
                ValueError(f"value failure from {surface}"),
            )
            for marker in markers:

                def fail_during_iteration(
                    _sql: str,
                    _parameters: tuple[object, ...],
                    _marker: BaseException = marker,
                ) -> Iterator[dict[str, object]]:
                    raise _marker
                    yield {}

                monkeypatch.setattr(graph, "_iter_rows", fail_during_iteration)
                if type(marker) is sqlite3.DatabaseError:
                    with pytest.raises(ExcelLSPError) as captured:
                        _invoke_graph_surface(graph, surface)
                    assert captured.value.code is ErrorCode.CORRUPT
                    assert captured.value.__cause__ is marker
                elif type(marker) is sqlite3.ProgrammingError:
                    with pytest.raises(sqlite3.ProgrammingError) as captured:
                        _invoke_graph_surface(graph, surface)
                    assert captured.value is marker
                else:
                    with pytest.raises(ValueError) as captured:
                        _invoke_graph_surface(graph, surface)
                    assert captured.value is marker

                assert store.connection.in_transaction is True
                monkeypatch.setattr(graph, "_iter_rows", original_iter_rows)
                assert _invoke_graph_surface(graph, surface) == expected
                assert store.connection.in_transaction is True

        if owner == "raw":
            store.connection.execute("BEGIN")
            try:
                exercise_errors()
            finally:
                store.connection.rollback()
        else:
            with store.transaction():
                exercise_errors()
                store.set_meta("snapshot_error_owner", surface)
            assert store.get_meta("snapshot_error_owner") == surface

        assert store.connection.in_transaction is False
        assert _invoke_graph_surface(graph, surface) == expected


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("cleanup_mode", ["fallback", "deny"])
@pytest.mark.parametrize("body_outcome", ["success", "error"])
@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_owned_snapshot_cleanup_failure_is_recovered_or_poisoned(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    journal_mode: str,
    cleanup_mode: str,
    body_outcome: str,
    surface: str,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    marker = ValueError(f"primary {surface} error")
    rollback_error = RuntimeError(f"rollback cleanup failed for {surface}")

    try:
        selected_mode = str(
            store.connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()[0]
        )
        assert selected_mode == journal_mode
        graph = DependencyGraph(store.connection)
        expected = _invoke_graph_surface(graph, surface)

        def fail_query_body() -> None:
            raise marker

        def fail_primary_rollback() -> None:
            raise rollback_error

        def deny_rollback(
            action: int,
            argument_one: str | None,
            argument_two: str | None,
            database: str | None,
            trigger: str | None,
        ) -> int:
            del argument_two, database, trigger
            if action == sqlite3.SQLITE_TRANSACTION and argument_one == "ROLLBACK":
                return sqlite3.SQLITE_DENY
            return sqlite3.SQLITE_OK

        with monkeypatch.context() as patch:
            if body_outcome == "error":
                patch.setattr(graph, "_require_clean_spatial", fail_query_body)
            if cleanup_mode == "fallback":
                patch.setattr(graph, "_rollback_owned_read_snapshot", fail_primary_rollback)
            else:
                store.connection.set_authorizer(deny_rollback)

            if body_outcome == "error":
                with pytest.raises(ValueError) as captured:
                    _invoke_graph_surface(graph, surface)
                assert captured.value is marker
                assert captured.value.__cause__ is not None
            else:
                with pytest.raises(BaseException) as captured:
                    _invoke_graph_surface(graph, surface)
                if cleanup_mode == "fallback":
                    assert captured.value is rollback_error
                else:
                    assert isinstance(captured.value, ExcelLSPError)
                    assert captured.value.code is ErrorCode.CORRUPT
                    assert isinstance(captured.value.__cause__, sqlite3.DatabaseError)
                    assert captured.value.__cause__.__cause__ is not None

        assert store.connection.in_transaction is False if cleanup_mode == "fallback" else True
    except sqlite3.ProgrammingError:
        assert cleanup_mode == "deny"

    writer = sqlite3.connect(store.path, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.execute(
            "INSERT INTO meta(key, value) VALUES ('snapshot_writer', 'unblocked') "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value"
        )
        writer.commit()
    finally:
        writer.close()

    if cleanup_mode == "fallback":
        # Any other-handle commit invalidates a cached facade conservatively;
        # a fresh facade revalidates the complete current sidecar.
        with pytest.raises(ExcelLSPError) as stale:
            _invoke_graph_surface(graph, surface)
        assert stale.value.code is ErrorCode.CORRUPT
        assert _invoke_graph_surface(DependencyGraph(store.connection), surface) == expected
        store.close()
    else:
        with pytest.raises(sqlite3.ProgrammingError):
            _invoke_graph_surface(graph, surface)
        store.close()
        with pytest.raises(RuntimeError, match="closed"):
            _ = store.connection


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("cleanup_mode", ["fallback", "deny"])
@pytest.mark.parametrize("body_chain", ["none", "cause", "context"])
def test_owned_snapshot_cleanup_preserves_body_causal_evidence(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    cleanup_mode: str,
    body_chain: str,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    graph = store.dependency_graph
    marker = ValueError("graph body primary")
    prior_error = LookupError(f"prior graph body {body_chain}")
    rollback_error = RuntimeError("injected graph rollback failure")

    def fail_query_body() -> None:
        if body_chain == "cause":
            raise marker from prior_error
        if body_chain == "context":
            try:
                raise prior_error
            except LookupError:
                raise marker  # noqa: B904 - exercise implicit exception context
        raise marker

    def fail_primary_rollback() -> None:
        raise rollback_error

    def deny_rollback(
        action: int,
        argument_one: str | None,
        argument_two: str | None,
        database: str | None,
        trigger: str | None,
    ) -> int:
        del argument_two, database, trigger
        if action == sqlite3.SQLITE_TRANSACTION and argument_one == "ROLLBACK":
            return sqlite3.SQLITE_DENY
        return sqlite3.SQLITE_OK

    with monkeypatch.context() as patch:
        patch.setattr(graph, "_require_clean_spatial", fail_query_body)
        if cleanup_mode == "fallback":
            patch.setattr(graph, "_rollback_owned_read_snapshot", fail_primary_rollback)
        else:
            store.connection.set_authorizer(deny_rollback)

        with pytest.raises(ValueError) as captured:
            _invoke_graph_surface(graph, "direct_precedents")

    assert captured.value is marker
    cleanup_cause = captured.value.__cause__
    assert cleanup_cause is not None
    if body_chain == "none":
        assert cleanup_cause is rollback_error or isinstance(cleanup_cause, BaseExceptionGroup)
    else:
        assert isinstance(cleanup_cause, BaseExceptionGroup)
        assert cleanup_cause.exceptions[0] is prior_error
        cleanup_cause = cleanup_cause.exceptions[1]

    if cleanup_mode == "fallback":
        assert cleanup_cause is rollback_error
        assert store.connection.in_transaction is False
    else:
        assert isinstance(cleanup_cause, BaseExceptionGroup)
        assert any(isinstance(error, sqlite3.DatabaseError) for error in cleanup_cause.exceptions)

    seen: set[int] = set()
    pending: list[BaseException] = [captured.value]
    while pending:
        error = pending.pop()
        assert id(error) not in seen
        seen.add(id(error))
        if error.__cause__ is not None:
            pending.append(error.__cause__)
        if error.__context__ is not None and not error.__suppress_context__:
            pending.append(error.__context__)
        if isinstance(error, BaseExceptionGroup):
            pending.extend(cast(tuple[BaseException, ...], error.exceptions))

    if cleanup_mode == "fallback":
        store.close()
    else:
        store.close()
        with pytest.raises(RuntimeError, match="closed"):
            _ = store.connection


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_poisoned_proxy_retries_close_until_finite_failure_releases_writer(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    graph = store.dependency_graph
    raw_connection = store.connection
    selected_mode = str(raw_connection.execute("PRAGMA journal_mode = DELETE").fetchone()[0])
    assert selected_mode == "delete"

    class FlakyCloseConnection:
        def __init__(self, connection: SQLiteConnectionLike) -> None:
            self.connection = connection
            self.close_calls = 0

        def close(self) -> None:
            self.close_calls += 1
            if self.close_calls <= 7:
                raise RuntimeError(f"transient close failure {self.close_calls}")
            self.connection.close()

        def __getattr__(self, name: str) -> Any:
            return getattr(self.connection, name)

    flaky_connection = FlakyCloseConnection(raw_connection)
    graph._connection = cast(sqlite3.Connection, flaky_connection)

    def deny_rollback(
        action: int,
        argument_one: str | None,
        argument_two: str | None,
        database: str | None,
        trigger: str | None,
    ) -> int:
        del argument_two, database, trigger
        if action == sqlite3.SQLITE_TRANSACTION and argument_one == "ROLLBACK":
            return sqlite3.SQLITE_DENY
        return sqlite3.SQLITE_OK

    raw_connection.set_authorizer(deny_rollback)
    with pytest.raises(ExcelLSPError) as first_failure:
        _invoke_graph_surface(graph, "direct_precedents")
    assert first_failure.value.code is ErrorCode.CORRUPT
    assert flaky_connection.close_calls == 4
    assert raw_connection.in_transaction is True

    with pytest.raises(sqlite3.ProgrammingError, match="poisoned"):
        _invoke_graph_surface(graph, "direct_precedents")
    assert flaky_connection.close_calls == 8
    with pytest.raises(sqlite3.ProgrammingError):
        _ = raw_connection.in_transaction

    writer = sqlite3.connect(store.path, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.execute(
            "INSERT INTO meta(key, value) VALUES ('finite_close_writer', 'unblocked') "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value"
        )
        writer.commit()
    finally:
        writer.close()

    store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_native_emergency_close_bypasses_long_subclass_failure_streak(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    database = store.path
    store.close()

    class FlakyNativeConnection(sqlite3.Connection):
        close_calls: int

        def close(self) -> None:
            self.close_calls += 1
            if self.close_calls <= 10:
                raise RuntimeError(f"native close-before failure {self.close_calls}")
            super().close()

    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=FlakyNativeConnection,
    )
    assert isinstance(connection, FlakyNativeConnection)
    connection.close_calls = 0
    selected_mode = str(connection.execute("PRAGMA journal_mode = DELETE").fetchone()[0])
    assert selected_mode == "delete"
    graph = DependencyGraph(connection, EdgeStore(connection))

    def deny_rollback(
        action: int,
        argument_one: str | None,
        argument_two: str | None,
        database_name: str | None,
        trigger: str | None,
    ) -> int:
        del argument_two, database_name, trigger
        if action == sqlite3.SQLITE_TRANSACTION and argument_one == "ROLLBACK":
            return sqlite3.SQLITE_DENY
        return sqlite3.SQLITE_OK

    connection.set_authorizer(deny_rollback)
    with pytest.raises(ExcelLSPError) as captured:
        _invoke_graph_surface(graph, "direct_precedents")
    assert captured.value.code is ErrorCode.CORRUPT
    assert connection.close_calls == 3
    with pytest.raises(sqlite3.ProgrammingError):
        _ = connection.in_transaction

    writer = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.execute(
            "INSERT INTO meta(key, value) VALUES ('native_emergency_writer', 'unblocked') "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value"
        )
        writer.commit()
    finally:
        writer.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("close_timing", ["before", "after"])
@pytest.mark.parametrize("body_outcome", ["success", "error"])
@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_poisoned_snapshot_retries_close_and_preserves_context_error_chain(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    journal_mode: str,
    close_timing: str,
    body_outcome: str,
    surface: str,
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    marker = ValueError(f"primary {surface} error")
    close_error = RuntimeError(f"close {close_timing} effect")
    selected_mode = str(
        store.connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()[0]
    )
    assert selected_mode == journal_mode
    graph = DependencyGraph(store.connection)

    def fail_query_body() -> None:
        raise marker

    def deny_rollback(
        action: int,
        argument_one: str | None,
        argument_two: str | None,
        database: str | None,
        trigger: str | None,
    ) -> int:
        del argument_two, database, trigger
        if action == sqlite3.SQLITE_TRANSACTION and argument_one == "ROLLBACK":
            return sqlite3.SQLITE_DENY
        return sqlite3.SQLITE_OK

    original_close = graph._close_poisoned_connection

    def failing_close() -> None:
        if close_timing == "after":
            original_close()
        raise close_error

    with monkeypatch.context() as patch:
        if body_outcome == "error":
            patch.setattr(graph, "_require_clean_spatial", fail_query_body)
        patch.setattr(graph, "_close_poisoned_connection", failing_close)
        store.connection.set_authorizer(deny_rollback)

        if body_outcome == "error":
            with pytest.raises(ValueError) as captured, store:
                _invoke_graph_surface(graph, surface)
            assert captured.value is marker
        else:
            with pytest.raises(ExcelLSPError) as captured, store:
                _invoke_graph_surface(graph, surface)
            assert captured.value.code is ErrorCode.CORRUPT

    cleanup_cause = captured.value.__cause__
    assert cleanup_cause is not None
    assert (
        close_error is cleanup_cause
        or (
            isinstance(cleanup_cause, BaseExceptionGroup)
            and close_error in cleanup_cause.exceptions
        )
        or (
            cleanup_cause.__cause__ is not None
            and isinstance(cleanup_cause.__cause__, BaseExceptionGroup)
            and close_error in cleanup_cause.__cause__.exceptions
        )
    )
    with pytest.raises(RuntimeError, match="closed"):
        _ = store.connection

    writer = sqlite3.connect(store.path, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.rollback()
    finally:
        writer.close()


@pytest.mark.parametrize(
    "surface",
    (
        "direct_precedents",
        "direct_dependents",
        "trace_precedents",
        "trace_dependents",
        "trace_path",
    ),
)
def test_public_query_boundary_preserves_input_value_errors(tmp_path: Path, surface: str) -> None:
    with _graph_store(tmp_path) as store:
        graph = DependencyGraph(store.connection, store.edge_store)
        opaque = GraphTarget("opaque", None, "opaque:test", None)

        with pytest.raises(ValueError):
            if surface == "direct_precedents":
                graph.direct_precedents(opaque)
            elif surface == "direct_dependents":
                graph.direct_dependents(opaque)
            elif surface == "trace_precedents":
                graph.trace_precedents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)), depth=9)
            elif surface == "trace_dependents":
                graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)), max_nodes=0)
            else:
                graph.trace_path(
                    GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
                    GraphArea(3, "Summary", Rect(1, 1, 3, 3)),
                    max_paths=0,
                )


@pytest.mark.parametrize("malformation", ["non-integer-epoch", "missing-state-table"])
def test_live_graph_trust_state_damage_is_shaped_as_corrupt(
    tmp_path: Path, malformation: str
) -> None:
    with _graph_store(tmp_path) as store:
        if malformation == "non-integer-epoch":
            store.connection.execute(
                "UPDATE graph_spatial_state SET mutation_epoch = 'invalid' WHERE singleton = 1"
            )
        else:
            store.connection.execute("DROP TABLE graph_spatial_state")

        with pytest.raises(ExcelLSPError, match="spatial mirrors") as captured:
            store.dependency_graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_direct_dependency_graph_construction_shapes_missing_state_as_corrupt(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        store.connection.execute("DROP TABLE graph_spatial_state")
        with pytest.raises(ExcelLSPError, match="storage schema") as captured:
            DependencyGraph(store.connection)
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("direction", ["dependents", "precedents"])
def test_cached_index_graph_rejects_live_rank_maximum_tampering(
    tmp_path: Path, prefer_rtree: bool, direction: str
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        graph = store.dependency_graph
        state_column = "dependent_rank_max" if direction == "dependents" else "precedent_rank_max"

        def query() -> TraceResult:
            if direction == "dependents":
                return graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
            return graph.trace_precedents(GraphArea(3, "Summary", Rect(1, 1, 3, 3)))

        assert query().node_count >= 2
        maximum = int(
            store.connection.execute(
                f"SELECT {state_column} FROM graph_spatial_state WHERE singleton = 1"
            ).fetchone()[0]
        )
        assert maximum >= 2
        store.connection.execute(
            f"UPDATE graph_spatial_state SET {state_column} = ? WHERE singleton = 1",
            (maximum - 1,),
        )

        with pytest.raises(ExcelLSPError, match="ranked spatial mirrors") as captured:
            query()
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    "mutation_sql",
    (
        "UPDATE graph_spatial_state SET dirty = 1 WHERE singleton = 1",
        "UPDATE graph_spatial_state SET dependent_rank_max = dependent_rank_max - 1 "
        "WHERE singleton = 1",
        "UPDATE graph_spatial_state SET precedent_rank_max = precedent_rank_max - 1 "
        "WHERE singleton = 1",
        "UPDATE graph_spatial_state SET revision = revision + 1 WHERE singleton = 1",
        "UPDATE graph_spatial_state SET mutation_epoch = mutation_epoch + 1 WHERE singleton = 1",
        "UPDATE graph_spatial_state SET clean_epoch = clean_epoch + 1 WHERE singleton = 1",
        "DELETE FROM graph_spatial_state WHERE singleton = 1",
    ),
    ids=(
        "dirty",
        "dependent-max",
        "precedent-max",
        "revision",
        "mutation-epoch",
        "clean-epoch",
        "missing-row",
    ),
)
def test_cached_index_graph_seals_every_mutable_trust_state_field(
    tmp_path: Path, prefer_rtree: bool, mutation_sql: str
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        graph = store.dependency_graph
        assert graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1))).node_count >= 2

        store.connection.execute(mutation_sql)

        with pytest.raises(ExcelLSPError, match="ranked spatial mirrors") as captured:
            graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("direction", ["dependents", "precedents"])
def test_lowered_persisted_rank_maximum_cannot_truncate_bounded_traces(
    tmp_path: Path, prefer_rtree: bool, direction: str
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    try:
        state_column = "dependent_rank_max" if direction == "dependents" else "precedent_rank_max"
        maximum = int(
            store.connection.execute(
                f"SELECT {state_column} FROM graph_spatial_state WHERE singleton = 1"
            ).fetchone()[0]
        )
        assert maximum >= 2
        store.set_meta("generation", "41")
        store.connection.execute(
            f"UPDATE graph_spatial_state SET {state_column} = ? WHERE singleton = 1",
            (maximum - 1,),
        )

        with pytest.raises(ExcelLSPError, match="storage schema") as captured:
            graph = DependencyGraph(store.connection)
            if direction == "dependents":
                graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
            else:
                graph.trace_precedents(GraphArea(3, "Summary", Rect(1, 1, 3, 3)))
        assert captured.value.code is ErrorCode.CORRUPT
        database = store.path
    finally:
        store.close()

    with IndexStore(database, prefer_rtree=prefer_rtree) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == 42


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("corruption", ["via-only", "coherent-rank-swap"])
def test_forged_clean_epoch_cannot_hide_noncanonical_semantic_ranks(
    tmp_path: Path, prefer_rtree: bool, corruption: str
) -> None:
    store = _graph_store(tmp_path, prefer_rtree=prefer_rtree)
    try:
        if corruption == "via-only":
            for edge_id, via in ((3, "opaque:alpha"), (4, "opaque:ZETA")):
                _insert_edge(
                    store,
                    edge_id=edge_id,
                    src_kind="fblock",
                    src_id=11,
                    src_sheet_id=2,
                    dst_sheet_id=None,
                    rect=None,
                    via=via,
                )
            # This reverses the exact casefold+case public ordering of edges 3
            # and 4 while leaving their persisted ranks and mirrors untouched.
            store.connection.execute("UPDATE edges SET via = 'opaque:zzzz' WHERE id = 3")
        else:
            # Swap all relational ranks and both matching mirrors coherently.
            # Density, maxima, geometry, and relational/mirror equality remain
            # valid; only the canonical public-hop order is false.
            store.connection.execute(
                """
                UPDATE edges
                SET dependent_rank = 3 - dependent_rank,
                    precedent_rank = 3 - precedent_rank
                """
            )
            for table in (
                store.edge_store.table_name,
                store.edge_store.source_table_name,
            ):
                if store.edge_store.backend == "rtree":
                    store.connection.execute(
                        f"""
                        UPDATE {table}
                        SET rank_min = 3 - rank_min,
                            rank_max = 3 - rank_max
                        """
                    )
                else:
                    store.connection.execute(f"UPDATE {table} SET rank = 3 - rank")

        store.set_meta("generation", "61")
        store.connection.execute(
            """
            UPDATE graph_spatial_state
            SET dirty = 0, clean_epoch = mutation_epoch
            WHERE singleton = 1
            """
        )

        with pytest.raises(EdgeSchemaError, match="canonical"):
            EdgeStore(store.connection)
        with pytest.raises(ExcelLSPError, match="storage schema") as captured:
            DependencyGraph(store.connection)
        assert captured.value.code is ErrorCode.CORRUPT
        database = store.path
    finally:
        store.close()

    # Invalid semantic sidecars are discarded through the normal monotonic
    # current-schema recovery path.
    with IndexStore(database, prefer_rtree=prefer_rtree) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == 62


def test_capped_precedents_match_direct_casefold_order_for_opaque_labels(
    tmp_path: Path,
) -> None:
    with _graph_store(tmp_path) as store:
        for edge_id, via in ((4, "opaque:ZETA"), (5, "opaque:alpha")):
            _insert_edge(
                store,
                edge_id=edge_id,
                src_kind="fblock",
                src_id=11,
                src_sheet_id=2,
                dst_sheet_id=None,
                rect=None,
                via=via,
            )
        graph = DependencyGraph(store.connection, store.edge_store)
        query = GraphArea(2, "Calc", Rect(1, 2, 2, 2))
        direct = graph.direct_precedents(query)
        capped = graph.trace_precedents(query, depth=1, max_nodes=3)

        assert [hop.target.label for hop in direct] == [
            "Inputs!A1:A1048576",
            "opaque:alpha",
            "opaque:ZETA",
        ]
        assert [child.target.label for child in capped.root.children] == [
            hop.target.label for hop in direct[:2]
        ]
        assert capped.truncated is True


@pytest.mark.parametrize("prefer_rtree", [True, False])
def test_capped_trace_uses_the_same_semantic_order_as_complete_direct_results(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
        for sheet_id, name in ((4, "Zulu"), (5, "Mike"), (6, "Alpha")):
            store.connection.execute(
                "INSERT INTO sheets VALUES (?, ?, ?, 'hash', 'worksheet', 'visible', 1, 2)",
                (sheet_id, name, f"xl/worksheets/sheet{sheet_id}.xml"),
            )
            _insert_fblock(
                store,
                block_id=20 + sheet_id,
                sheet_id=sheet_id,
                n=0,
                rect=Rect(1, 1, 2, 2),
            )
            _insert_edge(
                store,
                edge_id=10 + sheet_id,
                src_kind="fblock",
                src_id=20 + sheet_id,
                src_sheet_id=sheet_id,
                dst_sheet_id=1,
                rect=Rect(2, 2, 2, 2),
                via="ref",
            )
        graph = DependencyGraph(store.connection, store.edge_store)
        query = GraphArea(1, "Inputs", Rect(2, 2, 2, 2))

        direct = graph.direct_dependents(query)
        assert [hop.target.symbol for hop in direct] == [
            "fblock:Alpha:0",
            "fblock:Mike:0",
            "fblock:Zulu:0",
        ]
        capped = graph.trace_dependents(query, depth=1, max_nodes=2)
        assert capped.root.children[0].target.symbol == direct[0].target.symbol
        assert capped.truncated is True


@pytest.mark.parametrize("prefer_rtree", [True, False])
def test_semantic_prefix_is_stable_across_more_than_one_page_of_edge_id_churn(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    snapshots: list[tuple[tuple[str, ...], tuple[str, ...]]] = []
    for reverse_ids in (False, True):
        with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
            count = 1_200
            store.connection.executemany(
                "INSERT INTO fblocks VALUES (?, ?, ?, '=RC', ?, ?, 2, 2, 0, 0)",
                (
                    (10_000 + index, 2, 10_000 + index, 10 + index, 10 + index)
                    for index in range(count)
                ),
            )
            edge_ids = tuple(range(20_000, 20_000 + count))
            if reverse_ids:
                edge_ids = tuple(reversed(edge_ids))
            store.connection.executemany(
                """
                INSERT INTO edges(
                    id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES (?, 'fblock', ?, 2, 1, 2, 2, 2, 2, ?)
                """,
                (
                    (
                        edge_ids[index],
                        10_000 + index,
                        "name:ZETA" if index == 0 else "name:alpha",
                    )
                    for index in range(count)
                ),
            )
            store.rebuild_graph_spatial_index()
            graph = DependencyGraph(store.connection, store.edge_store)
            query = GraphArea(1, "Inputs", Rect(2, 2, 2, 2))
            direct = graph.direct_dependents(query)
            capped = graph.trace_dependents(query, depth=1, max_nodes=6)
            direct_labels = tuple(hop.target.label for hop in direct)
            capped_labels = tuple(child.target.label for child in capped.root.children)
            assert len(direct_labels) == count
            assert capped_labels == direct_labels[:5]
            assert capped.truncated is True

            exact_boundary = graph.trace_dependents(
                GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
                depth=1,
                max_nodes=2,
            )
            assert exact_boundary.root.child_count == 1
            assert exact_boundary.truncated is False
            snapshots.append((direct_labels, capped_labels))

    assert snapshots[0] == snapshots[1]


@pytest.mark.parametrize("prefer_rtree", [True, False])
def test_precedent_prefix_is_stable_across_more_than_one_page_of_edge_id_churn(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    snapshots: list[tuple[tuple[str, ...], tuple[str, ...]]] = []
    for reverse_ids in (False, True):
        with _graph_store(tmp_path, prefer_rtree=prefer_rtree) as store:
            count = 1_200
            edge_ids = tuple(range(30_000, 30_000 + count))
            if reverse_ids:
                edge_ids = tuple(reversed(edge_ids))
            store.connection.executemany(
                """
                INSERT INTO edges(
                    id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES (?, 'fblock', 11, 2, 1, ?, ?, 2, 2, ?)
                """,
                (
                    (
                        edge_ids[index],
                        10 + index,
                        10 + index,
                        "name:ZETA" if index == 0 else "name:alpha",
                    )
                    for index in range(count)
                ),
            )
            store.rebuild_graph_spatial_index()
            graph = store.dependency_graph
            query = GraphArea(2, "Calc", Rect(1, 2, 2, 2))
            direct = graph.direct_precedents(query)
            capped = graph.trace_precedents(query, depth=1, max_nodes=6)
            direct_labels = tuple(hop.target.label for hop in direct)
            capped_labels = tuple(child.target.label for child in capped.root.children)

            assert len(direct_labels) == count + 1
            assert capped_labels == direct_labels[:5]
            assert capped.truncated is True
            snapshots.append((direct_labels, capped_labels))

    assert snapshots[0] == snapshots[1]


def test_duplicate_edges_do_not_hide_later_unique_children_or_false_clear_cap(
    tmp_path: Path,
) -> None:
    with _graph_store(tmp_path, prefer_rtree=False) as store:
        for edge_id in range(10, 50):
            _insert_edge(
                store,
                edge_id=edge_id,
                src_kind="fblock",
                src_id=11,
                src_sheet_id=2,
                dst_sheet_id=1,
                rect=Rect(2, 2, 2, 2),
                via="ref",
            )
        _insert_edge(
            store,
            edge_id=50,
            src_kind="fblock",
            src_id=13,
            src_sheet_id=3,
            dst_sheet_id=1,
            rect=Rect(2, 2, 2, 2),
            via="ref",
        )
        result = DependencyGraph(store.connection, store.edge_store).trace_dependents(
            GraphArea(1, "Inputs", Rect(2, 2, 2, 2)), depth=1, max_nodes=3
        )

        assert [child.target.symbol for child in result.root.children] == [
            "fblock:Calc:0",
            "fblock:Summary:0",
        ]
        assert result.truncated is False


@pytest.mark.parametrize("source_kind", ["mystery", "fblock"])
def test_precedent_query_surfaces_unsupported_and_orphaned_sources(
    tmp_path: Path, source_kind: str
) -> None:
    with _graph_store(tmp_path) as store:
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (
                99, ?, 9999, 1, NULL, NULL, NULL, NULL, NULL, 'opaque:bad'
            )
            """,
            (source_kind,),
        )
        expected = "unsupported" if source_kind == "mystery" else "orphaned"
        with pytest.raises(ExcelLSPError, match=expected) as captured:
            store.rebuild_graph_spatial_index()
        assert captured.value.code is ErrorCode.CORRUPT


def test_path_truncation_detects_combinatorial_layered_dag_omissions(
    tmp_path: Path,
) -> None:
    with _graph_store(tmp_path) as store:
        _insert_fblock(store, block_id=14, sheet_id=2, n=1, rect=Rect(3, 3, 2, 2))
        _insert_fblock(store, block_id=15, sheet_id=3, n=1, rect=Rect(3, 3, 3, 3))
        _insert_fblock(store, block_id=16, sheet_id=3, n=2, rect=Rect(5, 5, 3, 3))
        edges = (
            (4, 14, 2, 1, Rect(1, 1, 1, 1)),
            (5, 13, 3, 2, Rect(3, 3, 2, 2)),
            (6, 15, 3, 2, Rect(1, 2, 2, 2)),
            (7, 15, 3, 2, Rect(3, 3, 2, 2)),
            (8, 16, 3, 3, Rect(1, 1, 3, 3)),
            (9, 16, 3, 3, Rect(3, 3, 3, 3)),
        )
        for edge_id, source_id, source_sheet, destination_sheet, rect in edges:
            _insert_edge(
                store,
                edge_id=edge_id,
                src_kind="fblock",
                src_id=source_id,
                src_sheet_id=source_sheet,
                dst_sheet_id=destination_sheet,
                rect=rect,
                via="ref",
            )
        result = DependencyGraph(store.connection, store.edge_store).trace_path(
            GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
            GraphArea(3, "Summary", Rect(5, 5, 3, 3)),
            max_paths=3,
        )

        assert result.connected is True
        assert len(result.paths) == 3
        assert result.truncated is True


@pytest.mark.parametrize(
    ("columns", "message"),
    [
        ((1, 1, None, 1), "partial rectangle"),
        ((1, 1, 1, 1), "orphaned"),
    ],
)
def test_corrupt_partial_destinations_and_orphan_sources_are_rejected(
    tmp_path: Path, columns: tuple[int | None, ...], message: str
) -> None:
    with _graph_store(tmp_path) as store:
        if message == "partial rectangle":
            store.connection.execute(
                """
                INSERT INTO edges(
                    id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES (
                    99, 'cell', 65537, 1, 1, ?, ?, ?, ?, 'ref'
                )
                """,
                columns,
            )
        else:
            store.connection.execute(
                """
                INSERT INTO edges(
                    id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES (
                    99, 'fblock', 9999, 1, 1, ?, ?, ?, ?, 'ref'
                )
                """,
                columns,
            )
        with pytest.raises(ExcelLSPError, match=message) as captured:
            store.rebuild_graph_spatial_index()
        assert captured.value.code is ErrorCode.CORRUPT


def test_graph_models_are_immutable() -> None:
    area = GraphArea(1, "My Sheet", Rect(1, 1, 1, 1))
    assert area.ref == "'My Sheet'!A1"
    with pytest.raises(FrozenInstanceError):
        area.sheet = "Changed"  # type: ignore[misc]


class _CommitObserver:
    """Expose a connection-compatible view that signals immediately before commit."""

    def __init__(
        self,
        connection: SQLiteConnectionLike,
        commit_attempted: Event,
        commit_returned: Event,
    ) -> None:
        self._connection = connection
        self._commit_attempted = commit_attempted
        self._commit_returned = commit_returned

    def __getattr__(self, name: str) -> Any:
        return getattr(self._connection, name)

    def commit(self) -> None:
        self._commit_attempted.set()
        self._connection.commit()
        self._commit_returned.set()


class _BeginFailureConnection(sqlite3.Connection):
    """Raise once immediately before or after a native graph BEGIN."""

    begin_timing: str | None = None
    begin_marker: sqlite3.DatabaseError

    def execute(self, sql: str, parameters: Any = (), /) -> sqlite3.Cursor:
        if self.begin_timing is not None and sql.strip().upper() == "BEGIN":
            timing = self.begin_timing
            self.begin_timing = None
            if timing == "after":
                super().execute(sql, parameters)
            raise self.begin_marker
        return super().execute(sql, parameters)


class _SpoofedNativeStateConnection(sqlite3.Connection):
    """Spoof virtual state while leaving the native descriptor observable."""

    spoof_after_begin: str | None = None
    state_spoof: str | None = None
    close_timing: str = "before"
    close_calls: int = 0

    @property
    def in_transaction(self) -> bool:
        if self.state_spoof == "false-positive":
            raise sqlite3.ProgrammingError("Cannot operate on a closed database.")
        if self.state_spoof == "false-negative":
            raise RuntimeError("virtual state probe failed after native close")
        descriptor = cast(Any, sqlite3.Connection.in_transaction)
        return bool(descriptor.__get__(self, sqlite3.Connection))

    def execute(self, sql: str, parameters: Any = (), /) -> sqlite3.Cursor:
        cursor = super().execute(sql, parameters)
        if sql.strip().upper() == "BEGIN" and self.spoof_after_begin is not None:
            self.state_spoof = self.spoof_after_begin
        return cursor

    def close(self) -> None:
        self.close_calls += 1
        if self.close_timing == "after":
            super().close()
        raise RuntimeError(f"virtual close {self.close_timing} native effect")


def _deny_graph_rollback(
    action: int,
    argument_one: str | None,
    argument_two: str | None,
    database_name: str | None,
    trigger: str | None,
) -> int:
    del argument_two, database_name, trigger
    if action == sqlite3.SQLITE_TRANSACTION and argument_one == "ROLLBACK":
        return sqlite3.SQLITE_DENY
    return sqlite3.SQLITE_OK


def _assert_unique_exception_identities(root: BaseException) -> None:
    seen: set[int] = set()
    pending = [root]
    while pending:
        error = pending.pop()
        assert id(error) not in seen
        seen.add(id(error))
        if error.__cause__ is not None:
            pending.append(error.__cause__)
        if error.__context__ is not None:
            pending.append(error.__context__)
        if isinstance(error, BaseExceptionGroup):
            pending.extend(cast(tuple[BaseException, ...], error.exceptions))


def _nested_cleanup_group(
    prior_error: BaseException,
) -> tuple[BaseExceptionGroup, BaseExceptionGroup, RuntimeError, RuntimeError]:
    distinct_one = RuntimeError("distinct nested cleanup one")
    distinct_two = RuntimeError("distinct nested cleanup two")
    distinct_two.__context__ = prior_error
    distinct_two.__suppress_context__ = True
    inner_group = BaseExceptionGroup(
        "inner nested cleanup",
        (prior_error, distinct_one),
    )
    outer_group = BaseExceptionGroup(
        "outer nested cleanup",
        (inner_group, distinct_two),
    )
    return outer_group, inner_group, distinct_one, distinct_two


def _graph_store(tmp_path: Path, *, prefer_rtree: bool = True) -> IndexStore:
    store = IndexStore(
        tmp_path / f"graph-{prefer_rtree}-{len(list(tmp_path.iterdir()))}.xlsp.db",
        prefer_rtree=prefer_rtree,
    )
    store.replace_sheet_catalog(
        (
            _sheet(1, "Inputs", 0),
            _sheet(2, "Calc", 1),
            _sheet(3, "Summary", 2),
        )
    )
    _insert_fblock(store, block_id=11, sheet_id=2, n=0, rect=Rect(1, 2, 2, 2))
    _insert_fblock(store, block_id=13, sheet_id=3, n=0, rect=Rect(1, 1, 3, 3))
    _insert_edge(
        store,
        edge_id=1,
        src_kind="fblock",
        src_id=11,
        src_sheet_id=2,
        dst_sheet_id=1,
        rect=Rect(1, 1_048_576, 1, 1),
        via="ref",
    )
    _insert_edge(
        store,
        edge_id=2,
        src_kind="fblock",
        src_id=13,
        src_sheet_id=3,
        dst_sheet_id=2,
        rect=Rect(1, 2, 2, 2),
        via="ref",
    )
    return store


def _sheet(sheet_id: int, name: str, order: int) -> SheetDescriptor:
    return SheetDescriptor(
        order=order,
        name=name,
        sheet_id=sheet_id,
        rel_id=f"rId{sheet_id}",
        xml_part=f"xl/worksheets/sheet{sheet_id}.xml",
        kind="worksheet",
    )


def _insert_fblock(store: IndexStore, *, block_id: int, sheet_id: int, n: int, rect: Rect) -> None:
    store.connection.execute(
        """
        INSERT INTO fblocks VALUES (?, ?, ?, '=RC', ?, ?, ?, ?, 0, 0)
        """,
        (
            block_id,
            sheet_id,
            n,
            rect.row_min,
            rect.row_max,
            rect.col_min,
            rect.col_max,
        ),
    )


def _insert_edge(
    store: IndexStore,
    *,
    edge_id: int,
    src_kind: str,
    src_id: int,
    src_sheet_id: int,
    dst_sheet_id: int | None,
    rect: Rect | None,
    via: str,
) -> None:
    store.connection.execute(
        """
        INSERT INTO edges(
            id, src_kind, src_id, src_sheet_id, dst_sheet_id,
            dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            edge_id,
            src_kind,
            src_id,
            src_sheet_id,
            dst_sheet_id,
            None if rect is None else rect.row_min,
            None if rect is None else rect.row_max,
            None if rect is None else rect.col_min,
            None if rect is None else rect.col_max,
            via,
        ),
    )
    store.rebuild_graph_spatial_index()


def _invoke_graph_surface(graph: DependencyGraph, surface: str) -> object:
    if surface == "direct_precedents":
        return graph.direct_precedents(GraphArea(2, "Calc", Rect(1, 2, 2, 2)))
    if surface == "direct_dependents":
        return graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
    if surface == "trace_precedents":
        return graph.trace_precedents(GraphArea(3, "Summary", Rect(1, 1, 3, 3)))
    if surface == "trace_dependents":
        return graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
    if surface == "trace_path":
        return graph.trace_path(
            GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
            GraphArea(3, "Summary", Rect(1, 1, 3, 3)),
        )
    raise AssertionError(f"unknown graph surface {surface!r}")
