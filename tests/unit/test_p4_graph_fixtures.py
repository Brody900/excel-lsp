from __future__ import annotations

import json
import os
import runpy
from collections import Counter
from collections.abc import Callable, Iterable, Sequence
from pathlib import Path
from typing import cast

import pytest
import tiktoken

import excel_lsp.core.graph.circular as circular_module
import excel_lsp.core.index.lifecycle as lifecycle_module
import excel_lsp.core.index.store as store_module
from excel_lsp.core.graph import GraphArea, GraphHop, PathResult, TraceNode, TraceResult
from excel_lsp.core.graph.circular import (
    BlockKey,
    CellDependency,
    CellNode,
    CircularBlock,
)
from excel_lsp.core.index import IndexStore, index_workbook
from excel_lsp.core.models import (
    DefinedName,
    NameArea,
    Rect,
    SheetDescriptor,
    WorkbookMetadata,
)

GenerateAll = Callable[[Path], dict[str, Path]]
generate_all = cast(
    GenerateAll,
    runpy.run_path(str(Path(__file__).parents[1] / "fixtures" / "generate.py"))["generate_all"],
)

_GRAPH_FIXTURES = ("F03", "F04", "F05", "F09a", "F09b", "F15", "F19")
_UPDATE_GOLDEN_ENV = "EXCEL_LSP_UPDATE_P4_GOLDENS"
_P4_GOLDEN_TOKEN_CAP = 4_000
_F03_EXPECTED_EDGE_PROJECTION = (
    # Calc!B2 = Inputs!$B$3
    ("fblock", "Calc", 2, 2, 2, 2, "Inputs", 3, 3, 2, 2, "ref"),
    # Calc!B3:B6 = B(previous row) * (1 + Inputs!$B$2)
    ("fblock", "Calc", 3, 6, 2, 2, "Calc", 2, 5, 2, 2, "ref"),
    ("fblock", "Calc", 3, 6, 2, 2, "Inputs", 2, 2, 2, 2, "ref"),
    # Calc!C2:C6 = B(same row) * Inputs!$B$4
    ("fblock", "Calc", 2, 6, 3, 3, "Calc", 2, 6, 2, 2, "ref"),
    ("fblock", "Calc", 2, 6, 3, 3, "Inputs", 4, 4, 2, 2, "ref"),
    # Calc!D2:D6 = (B(same row) - C(same row)) * (1 - Inputs!$B$5)
    ("fblock", "Calc", 2, 6, 4, 4, "Calc", 2, 6, 2, 2, "ref"),
    ("fblock", "Calc", 2, 6, 4, 4, "Calc", 2, 6, 3, 3, "ref"),
    ("fblock", "Calc", 2, 6, 4, 4, "Inputs", 5, 5, 2, 2, "ref"),
    # Summary annual values and KPI singletons.
    ("fblock", "Summary", 2, 6, 3, 3, "Calc", 2, 6, 4, 4, "ref"),
    ("fblock", "Summary", 7, 7, 3, 3, "Calc", 2, 2, 2, 2, "ref"),
    ("fblock", "Summary", 8, 8, 3, 3, "Calc", 6, 6, 2, 2, "ref"),
    ("fblock", "Summary", 9, 9, 3, 3, "Calc", 2, 6, 3, 3, "ref"),
    ("fblock", "Summary", 10, 10, 3, 3, "Calc", 2, 6, 4, 4, "ref"),
)
_F04_EXPECTED_EDGE_PROJECTION = (
    # Inputs!B4 = BaseAmount * (1 + GlobalRate)
    ("fblock", "Inputs", 4, 4, 2, 2, "Inputs", 2, 2, 2, 2, "name:BaseAmount"),
    ("fblock", "Inputs", 4, 4, 2, 2, "Inputs", 3, 3, 2, 2, "name:GlobalRate"),
    # Calc!B3 = BaseAmount
    ("fblock", "Calc", 3, 3, 2, 2, "Inputs", 2, 2, 2, 2, "name:BaseAmount"),
    # Calc!B4 = BaseAmount * (1 + ScopedRate)
    ("fblock", "Calc", 4, 4, 2, 2, "Inputs", 2, 2, 2, 2, "name:BaseAmount"),
    ("fblock", "Calc", 4, 4, 2, 2, "Calc", 2, 2, 2, 2, "name:ScopedRate"),
)
_F05_EXPECTED_EDGE_PROJECTION = (
    # Structured!D2:D5 = [@Qty] * [@Price]
    (
        "fblock",
        "Structured",
        2,
        5,
        4,
        4,
        "Structured",
        2,
        5,
        2,
        2,
        "structured:Table1[Qty]",
    ),
    (
        "fblock",
        "Structured",
        2,
        5,
        4,
        4,
        "Structured",
        2,
        5,
        3,
        3,
        "structured:Table1[Price]",
    ),
    # Structured!B6 = SUBTOTAL(109, Table1[Qty])
    (
        "fblock",
        "Structured",
        6,
        6,
        2,
        2,
        "Structured",
        2,
        5,
        2,
        2,
        "structured:Table1[Qty]",
    ),
    # Structured!D6 = SUBTOTAL(109, Table1[LineTotal])
    (
        "fblock",
        "Structured",
        6,
        6,
        4,
        4,
        "Structured",
        2,
        5,
        4,
        4,
        "structured:Table1[LineTotal]",
    ),
    # Structured!F2 = SUM(Table1[LineTotal])
    (
        "fblock",
        "Structured",
        2,
        2,
        6,
        6,
        "Structured",
        2,
        5,
        4,
        4,
        "structured:Table1[LineTotal]",
    ),
)
_F15_EXPECTED_EDGE_PROJECTION = (
    # Summary!B2 = SUM(Jan:Mar!B2), expanded in workbook sheet order.
    ("fblock", "Summary", 2, 2, 2, 2, "Jan", 2, 2, 2, 2, "3d"),
    ("fblock", "Summary", 2, 2, 2, 2, "Feb", 2, 2, 2, 2, "3d"),
    ("fblock", "Summary", 2, 2, 2, 2, "Mar", 2, 2, 2, 2, "3d"),
)
_F19_EXPECTED_EDGE_PROJECTION = (
    # FILTER's two identical I2:I4 tokens coalesce to one semantic edge.
    ("fblock", "Modern", 1, 1, 1, 1, "Modern", 2, 4, 9, 9, "ref"),
    # Spill anchors.
    ("fblock", "Modern", 1, 1, 2, 2, "Modern", 1, 1, 1, 1, "spill"),
    ("fblock", "Modern", 1, 1, 3, 3, "Modern", 1, 1, 1, 1, "spill"),
    # LET binding value, LAMBDA invocation argument, XLOOKUP arrays, and @ range.
    ("fblock", "Modern", 1, 1, 4, 4, "Modern", 2, 2, 9, 9, "ref"),
    ("fblock", "Modern", 1, 1, 5, 5, "Modern", 3, 3, 9, 9, "ref"),
    ("fblock", "Modern", 1, 1, 6, 6, "Modern", 2, 4, 8, 8, "ref"),
    ("fblock", "Modern", 1, 1, 6, 6, "Modern", 2, 4, 9, 9, "ref"),
    ("fblock", "Modern", 2, 2, 7, 7, "Modern", 2, 4, 9, 9, "ref"),
)
_EXPECTED_EDGE_PROJECTIONS = {
    "F03": _F03_EXPECTED_EDGE_PROJECTION,
    "F04": _F04_EXPECTED_EDGE_PROJECTION,
    "F05": _F05_EXPECTED_EDGE_PROJECTION,
    "F15": _F15_EXPECTED_EDGE_PROJECTION,
    "F19": _F19_EXPECTED_EDGE_PROJECTION,
}
_TRACE_CASES = {
    "F03": (
        GraphArea(1, "Inputs", Rect(2, 2, 2, 2)),
        GraphArea(3, "Summary", Rect(10, 10, 3, 3)),
    ),
    "F04": (
        GraphArea(1, "Inputs", Rect(2, 2, 2, 2)),
        GraphArea(2, "Calc", Rect(4, 4, 2, 2)),
    ),
    "F05": (
        GraphArea(1, "Structured", Rect(2, 2, 2, 2)),
        GraphArea(1, "Structured", Rect(2, 2, 6, 6)),
    ),
    "F15": (
        GraphArea(1, "Jan", Rect(2, 2, 2, 2)),
        GraphArea(4, "Summary", Rect(2, 2, 2, 2)),
    ),
    "F19": (
        GraphArea(1, "Modern", Rect(2, 2, 9, 9)),
        GraphArea(1, "Modern", Rect(1, 1, 1, 1)),
    ),
}


@pytest.fixture(scope="module")
def indexed_fixtures(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    root = tmp_path_factory.mktemp("p4-graph-fixtures")
    fixture_paths = generate_all(root / "workbooks")
    index_dir = root / "indexes"
    return {
        fixture_id: Path(index_workbook(fixture_paths[fixture_id], index_dir=index_dir).index_path)
        for fixture_id in _GRAPH_FIXTURES
    }


def _edge_projection(store: IndexStore) -> tuple[tuple[object, ...], ...]:
    """Project all edge sources by geometry, including unexpected cell sources."""
    return tuple(
        tuple(row)
        for row in store.connection.execute(
            """
            SELECT e.src_kind, src.name,
                   CASE WHEN e.src_kind = 'fblock' THEN fb.row_min
                        WHEN e.src_kind = 'cell' THEN e.src_id >> 16 END,
                   CASE WHEN e.src_kind = 'fblock' THEN fb.row_max
                        WHEN e.src_kind = 'cell' THEN e.src_id >> 16 END,
                   CASE WHEN e.src_kind = 'fblock' THEN fb.col_min
                        WHEN e.src_kind = 'cell' THEN e.src_id & 65535 END,
                   CASE WHEN e.src_kind = 'fblock' THEN fb.col_max
                        WHEN e.src_kind = 'cell' THEN e.src_id & 65535 END,
                   dst.name, e.dst_row_min, e.dst_row_max,
                   e.dst_col_min, e.dst_col_max, e.via
            FROM edges AS e
            LEFT JOIN sheets AS src ON src.id = e.src_sheet_id
            LEFT JOIN fblocks AS fb
                   ON e.src_kind = 'fblock' AND fb.id = e.src_id
            LEFT JOIN sheets AS dst ON dst.id = e.dst_sheet_id
            ORDER BY e.id
            """
        )
    )


def _expected_destination_area(
    edge: tuple[object, ...],
    sheet_ids: dict[str, int],
) -> GraphArea:
    destination_sheet = cast(str, edge[6])
    return GraphArea(
        sheet_ids[destination_sheet],
        destination_sheet,
        Rect(
            cast(int, edge[7]),
            cast(int, edge[8]),
            cast(int, edge[9]),
            cast(int, edge[10]),
        ),
    )


def _expected_precedent_hop_projection(
    edge: tuple[object, ...],
    sheet_ids: dict[str, int],
) -> tuple[str, str | None, str, GraphArea, str]:
    area = _expected_destination_area(edge, sheet_ids)
    is_cell = area.rect.row_min == area.rect.row_max and area.rect.col_min == area.rect.col_max
    symbol = f"cell:{area.sheet}!{area.ref.rsplit('!', 1)[-1]}" if is_cell else None
    return (
        "cell" if is_cell else "range",
        symbol,
        area.ref,
        area,
        cast(str, edge[11]),
    )


def _public_hop_projection(hop: GraphHop) -> tuple[str, str | None, str, GraphArea, str]:
    assert hop.target.area is not None
    assert hop.target.ref is not None
    return (
        hop.target.kind,
        hop.target.symbol,
        hop.target.ref,
        hop.target.area,
        hop.via,
    )


def test_f03_edge_projection_is_exact(indexed_fixtures: dict[str, Path]) -> None:
    with IndexStore(indexed_fixtures["F03"]) as store:
        actual = _edge_projection(store)

    assert Counter(actual) == Counter(_F03_EXPECTED_EDGE_PROJECTION)
    assert (
        "fblock",
        "Summary",
        7,
        7,
        3,
        3,
        "Calc",
        2,
        2,
        2,
        2,
        "ref",
    ) in actual


@pytest.mark.parametrize(
    ("fixture_id", "expected"),
    (
        ("F04", _F04_EXPECTED_EDGE_PROJECTION),
        ("F05", _F05_EXPECTED_EDGE_PROJECTION),
        ("F15", _F15_EXPECTED_EDGE_PROJECTION),
        ("F19", _F19_EXPECTED_EDGE_PROJECTION),
    ),
)
def test_p4_fixture_edges_are_exact(
    indexed_fixtures: dict[str, Path],
    fixture_id: str,
    expected: tuple[tuple[object, ...], ...],
) -> None:
    with IndexStore(indexed_fixtures[fixture_id]) as store:
        actual = _edge_projection(store)

    assert Counter(actual) == Counter(expected)


def test_edge_projection_rejects_duplicate_and_unexpected_valid_edges(
    indexed_fixtures: dict[str, Path],
) -> None:
    expected = Counter(_F04_EXPECTED_EDGE_PROJECTION)
    with IndexStore(indexed_fixtures["F04"]) as store:
        assert Counter(_edge_projection(store)) == expected

        store.connection.execute("SAVEPOINT duplicate_required_edge")
        try:
            store.connection.execute(
                """
                INSERT INTO edges(
                    src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via,
                    dependent_rank, precedent_rank
                )
                SELECT src_kind, src_id, src_sheet_id, dst_sheet_id,
                       dst_row_min, dst_row_max, dst_col_min, dst_col_max, via,
                       dependent_rank, precedent_rank
                FROM edges ORDER BY id LIMIT 1
                """
            )
            assert Counter(_edge_projection(store)) != expected
        finally:
            store.connection.execute("ROLLBACK TO duplicate_required_edge")
            store.connection.execute("RELEASE duplicate_required_edge")

        store.connection.execute("SAVEPOINT unexpected_singleton_edge")
        try:
            source_sheet_id = int(
                store.connection.execute("SELECT id FROM sheets WHERE name = 'Inputs'").fetchone()[
                    0
                ]
            )
            destination_sheet_id = int(
                store.connection.execute("SELECT id FROM sheets WHERE name = 'Calc'").fetchone()[0]
            )
            store.connection.execute(
                """
                INSERT INTO edges(
                    src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES ('cell', ?, ?, ?, 1, 1, 1, 1, 'ref')
                """,
                ((1 << 16) | 1, source_sheet_id, destination_sheet_id),
            )
            assert Counter(_edge_projection(store)) != expected
        finally:
            store.connection.execute("ROLLBACK TO unexpected_singleton_edge")
            store.connection.execute("RELEASE unexpected_singleton_edge")


def test_schema_v5_has_bounded_graph_traversal_indexes(
    indexed_fixtures: dict[str, Path],
) -> None:
    with IndexStore(indexed_fixtures["F03"]) as store:
        indexes = {str(row[1]) for row in store.connection.execute("PRAGMA index_list(edges)")}
        assert {"edges_source", "edges_precedent_semantic"} <= indexes


@pytest.mark.parametrize("prefer_rtree", (True, False), ids=("rtree", "interval"))
@pytest.mark.parametrize("fixture_id", tuple(_EXPECTED_EDGE_PROJECTIONS))
def test_i12_every_concrete_precedent_round_trips_to_its_source_block(
    indexed_fixtures: dict[str, Path],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    fixture_id: str,
    prefer_rtree: bool,
) -> None:
    with IndexStore(indexed_fixtures[fixture_id]) as source_store:
        workbook_path = source_store.get_meta("workbook_path")
    assert workbook_path is not None

    with monkeypatch.context() as patch:
        patch.setattr(
            lifecycle_module,
            "IndexStore",
            lambda path: IndexStore(path, prefer_rtree=prefer_rtree),
        )
        update = index_workbook(
            Path(workbook_path),
            index_dir=tmp_path / f"i12-{fixture_id}-{prefer_rtree}",
        )

    expected_edges = _EXPECTED_EDGE_PROJECTIONS[fixture_id]
    expected_by_source: dict[
        tuple[str, int, int, int, int],
        list[tuple[object, ...]],
    ] = {}
    for edge in expected_edges:
        source_key = (
            cast(str, edge[1]),
            cast(int, edge[2]),
            cast(int, edge[3]),
            cast(int, edge[4]),
            cast(int, edge[5]),
        )
        expected_by_source.setdefault(source_key, []).append(edge)

    with IndexStore(update.index_path, prefer_rtree=prefer_rtree) as store:
        assert store.edge_store.backend == ("rtree" if prefer_rtree else "interval")
        graph = store.dependency_graph
        sheet_ids = {
            str(row["name"]): int(row["id"])
            for row in store.connection.execute("SELECT id, name FROM sheets")
        }
        checked = 0
        for source_key, source_edges in expected_by_source.items():
            source_sheet, row_min, row_max, col_min, col_max = source_key
            block_rows = store.connection.execute(
                """
                SELECT f.n
                FROM fblocks AS f JOIN sheets AS s ON s.id = f.sheet_id
                WHERE s.name = ? AND f.row_min = ? AND f.row_max = ?
                  AND f.col_min = ? AND f.col_max = ?
                """,
                source_key,
            ).fetchall()
            assert len(block_rows) == 1, (fixture_id, source_key, block_rows)
            source = GraphArea(
                sheet_ids[source_sheet],
                source_sheet,
                Rect(row_min, row_max, col_min, col_max),
            )
            source_symbol = f"fblock:{source_sheet}:{block_rows[0]['n']}"
            expected_hops = tuple(
                _expected_precedent_hop_projection(edge, sheet_ids) for edge in source_edges
            )
            actual_hops = tuple(
                _public_hop_projection(hop) for hop in graph.direct_precedents(source)
            )
            assert Counter(actual_hops) == Counter(expected_hops), (fixture_id, source_key)

            for edge in source_edges:
                destination = _expected_destination_area(edge, sheet_ids)
                via = cast(str, edge[11])
                dependent_hops = {
                    (hop.target.symbol, hop.via) for hop in graph.direct_dependents(destination)
                }
                assert (source_symbol, via) in dependent_hops, (
                    fixture_id,
                    source_key,
                    destination,
                    via,
                )
                checked += 1
        assert checked == len(expected_edges)


def test_f09a_indexes_to_exactly_one_circular_error_with_path(
    indexed_fixtures: dict[str, Path],
) -> None:
    with IndexStore(indexed_fixtures["F09a"]) as store:
        rows = store.connection.execute(
            """
            SELECT severity, code, ref, related FROM diagnostics
            WHERE code IN ('E_CIRCULAR', 'W_POSSIBLE_CIRCULAR')
            ORDER BY code, ref
            """
        ).fetchall()
        assert len(rows) == 1
        assert tuple(rows[0][:3]) == ("error", "E_CIRCULAR", "cell:Circular!B2")
        related = json.loads(str(rows[0]["related"]))
        assert related == {
            "candidate_blocks": ["fblock:Circular:0", "fblock:Circular:1"],
            "path": ["cell:Circular!B2", "cell:Circular!B3", "cell:Circular!B2"],
        }


def test_f09b_exercises_coarse_self_overlap_but_stays_clean(
    indexed_fixtures: dict[str, Path],
) -> None:
    with IndexStore(indexed_fixtures["F09b"]) as store:
        assert tuple(
            map(
                tuple,
                store.connection.execute(
                    "SELECT n, row_min, row_max, col_min, col_max FROM fblocks"
                ),
            )
        ) == ((0, 3, 50_002, 2, 2),)
        assert tuple(
            map(
                tuple,
                store.connection.execute(
                    """
                    SELECT dst_row_min, dst_row_max, dst_col_min, dst_col_max
                    FROM edges
                    """
                ),
            )
        ) == ((2, 50_001, 2, 2),)
        assert (
            store.connection.execute(
                """
                SELECT count(*) FROM diagnostics
                WHERE code IN ('E_CIRCULAR', 'W_POSSIBLE_CIRCULAR')
                """
            ).fetchone()[0]
            == 0
        )


def test_i13_real_f09b_never_allocates_the_stage_2b_exact_graph(
    indexed_fixtures: dict[str, Path],
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    with IndexStore(indexed_fixtures["F09b"]) as store:
        workbook_path = store.get_meta("workbook_path")
    assert workbook_path is not None

    resolved_cells: list[CellNode] = []
    stage_2a_results: list[tuple[int, CellNode | None, bool, bool]] = []
    original_self_check = circular_module._check_self_inclusion

    def observe_stage_2a(
        blocks: Sequence[CircularBlock],
        resolve_exact: Callable[[CellNode], Iterable[CellDependency]],
    ) -> object:
        def counting_resolver(cell: CellNode) -> Iterable[CellDependency]:
            resolved_cells.append(cell)
            return resolve_exact(cell)

        result = original_self_check(blocks, counting_resolver)
        stage_2a_results.append(
            (len(blocks), result.hit, result.all_internal_before, result.all_internal_after)
        )
        return result

    def reject_stage_2b(*_args: object, **_kwargs: object) -> None:
        pytest.fail("F09b entered stage 2b and attempted to allocate an exact cell graph")

    monkeypatch.setattr(circular_module, "_check_self_inclusion", observe_stage_2a)
    monkeypatch.setattr(circular_module, "_expand_exact_graph", reject_stage_2b)
    update = index_workbook(Path(workbook_path), index_dir=tmp_path / "indexes")

    assert stage_2a_results == [(1, None, True, False)]
    expected_cells = [CellNode(1, row, 2) for row in range(3, 50_003)]
    assert len(resolved_cells) == len(expected_cells) == 50_000
    assert resolved_cells == expected_cells
    with IndexStore(update.index_path) as store:
        assert (
            store.connection.execute(
                """
                SELECT count(*) FROM diagnostics
                WHERE code IN ('E_CIRCULAR', 'W_POSSIBLE_CIRCULAR')
                """
            ).fetchone()[0]
            == 0
        )


def test_f03_trace_and_path_reach_summary_total(
    indexed_fixtures: dict[str, Path],
) -> None:
    with IndexStore(indexed_fixtures["F03"]) as store:
        graph = store.dependency_graph
        source = GraphArea(1, "Inputs", Rect(2, 2, 2, 2))
        destination = GraphArea(3, "Summary", Rect(10, 10, 3, 3))

        trace = graph.trace_dependents(source, depth=8, max_nodes=200)
        assert trace.truncated is False
        assert "fblock:Summary:4" in _trace_symbols(trace.root)

        paths = graph.trace_path(source, destination, max_paths=3, max_depth=12)
        assert paths.connected is True
        assert paths.paths
        assert all(path[0].symbol == "cell:Inputs!B2" for path in paths.paths)
        assert all(path[-1].symbol == "fblock:Summary:4" for path in paths.paths)


def test_p4_trace_snapshot_matches_or_updates_golden(
    indexed_fixtures: dict[str, Path],
) -> None:
    cases: dict[str, object] = {}
    for fixture_id, (source, destination) in _TRACE_CASES.items():
        with IndexStore(indexed_fixtures[fixture_id]) as store:
            graph = store.dependency_graph
            cases[fixture_id] = {
                "dependents": _trace_snapshot(
                    graph.trace_dependents(source, depth=3, max_nodes=50)
                ),
                "precedents": _trace_snapshot(
                    graph.trace_precedents(destination, depth=3, max_nodes=50)
                ),
                "path": _path_snapshot(
                    graph.trace_path(source, destination, max_paths=3, max_depth=12)
                ),
            }

    circular: dict[str, object] = {}
    for fixture_id in ("F09a", "F09b"):
        with IndexStore(indexed_fixtures[fixture_id]) as store:
            circular[fixture_id] = [
                list(row)
                for row in store.connection.execute(
                    """
                    SELECT severity, code, s.name, d.row, d.col, d.ref,
                           d.message, d.related
                    FROM diagnostics AS d JOIN sheets AS s ON s.id = d.sheet_id
                    WHERE d.code IN ('E_CIRCULAR', 'W_POSSIBLE_CIRCULAR')
                    ORDER BY s.id, d.row, d.col, d.code
                    """
                )
            ]
    serialized = (
        json.dumps(
            {"traces": cases, "circular": circular},
            ensure_ascii=False,
            indent=2,
        )
        + "\n"
    )
    token_count = len(tiktoken.get_encoding("o200k_base").encode(serialized))
    assert token_count <= _P4_GOLDEN_TOKEN_CAP
    golden_path = Path(__file__).parents[1] / "golden" / "p4-graph-semantics.json"
    if os.environ.get(_UPDATE_GOLDEN_ENV) == "1":
        golden_path.write_text(serialized, encoding="utf-8", newline="\n")
    assert golden_path.read_bytes() == serialized.encode("utf-8")


def test_circular_adapter_keeps_equal_packed_cell_ids_sheet_local(tmp_path: Path) -> None:
    descriptors = tuple(
        SheetDescriptor(
            order=order,
            name=name,
            sheet_id=order + 1,
            rel_id=f"rId{order + 1}",
            xml_part=f"xl/worksheets/sheet{order + 1}.xml",
            kind="worksheet",
        )
        for order, name in enumerate(("One", "Two"))
    )
    metadata = WorkbookMetadata(
        path=str(tmp_path / "cell-sources.xlsx"),
        date1904=False,
        sheets=descriptors,
        defined_names=(),
    )
    packed_a1 = (1 << 16) | 1
    with IndexStore(tmp_path / "cell-sources.xlsp.db") as store:
        store.replace_sheet_catalog(descriptors)
        for sheet_id in (1, 2):
            store.connection.execute(
                """
                INSERT INTO cells(
                    sheet_id, row, col, ref, value, value_type, formula,
                    style_idx, formula_kind, shared_index, array_ref, data_table
                ) VALUES (?, 1, 1, 'A1', 0, 'number', '=A1', 0,
                          'normal', NULL, NULL, NULL)
                """,
                (sheet_id,),
            )
            cursor = store.connection.execute(
                """
                INSERT INTO edges(
                    src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES ('cell', ?, ?, ?, 1, 1, 1, 1, 'ref')
                """,
                (packed_a1, sheet_id, sheet_id),
            )
            assert cursor.lastrowid is not None
            store.edge_store.insert(int(cursor.lastrowid), sheet_id, Rect(1, 1, 1, 1))

        context = store._formula_reference_context(metadata)
        store._replace_circular_diagnostics(metadata, context)

        assert tuple(
            map(
                tuple,
                store.connection.execute(
                    """
                    SELECT code, s.name, d.ref
                    FROM diagnostics AS d JOIN sheets AS s ON s.id = d.sheet_id
                    WHERE code = 'E_CIRCULAR'
                    ORDER BY s.id
                    """
                ),
            )
        ) == (
            ("E_CIRCULAR", "One", "cell:One!A1"),
            ("E_CIRCULAR", "Two", "cell:Two!A1"),
        )


def test_circular_owner_index_avoids_all_block_scan(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    blocks = tuple(
        CircularBlock(BlockKey(1, col - 1), Rect(1, 1_048_576, col, col)) for col in range(1, 4097)
    )
    index = store_module._build_circular_owner_index(blocks)
    original = store_module._query_circular_owner_index
    calls = 0

    def observed(
        node: store_module._CircularOwnerNode | None,
        row: int,
        col: int,
    ) -> tuple[CircularBlock, ...]:
        nonlocal calls
        calls += 1
        return original(node, row, col)

    monkeypatch.setattr(store_module, "_query_circular_owner_index", observed)

    assert observed(index, 500_000, 2048) == (blocks[2047],)
    assert calls < 64


def test_circular_adapter_preserves_coordinate_spill_anchor(tmp_path: Path) -> None:
    descriptor = SheetDescriptor(
        order=0,
        name="Spill",
        sheet_id=1,
        rel_id="rId1",
        xml_part="xl/worksheets/sheet1.xml",
        kind="worksheet",
    )
    metadata = WorkbookMetadata(
        path=str(tmp_path / "spill.xlsx"),
        date1904=False,
        sheets=(descriptor,),
        defined_names=(),
    )
    with IndexStore(tmp_path / "spill.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))
        for row in (2, 3):
            store.connection.execute(
                """
                INSERT INTO cells(
                    sheet_id, row, col, ref, value, value_type, formula,
                    style_idx, formula_kind, shared_index, array_ref, data_table
                ) VALUES (1, ?, 2, ?, 0, 'number', '=B3#', 0,
                          'normal', NULL, NULL, NULL)
                """,
                (row, f"B{row}"),
            )
        cursor = store.connection.execute(
            """
            INSERT INTO fblocks(
                sheet_id, n, r1c1, row_min, row_max, col_min, col_max,
                volatile, opaque
            ) VALUES (1, 0, '=B3#', 2, 3, 2, 2, 0, 0)
            """
        )
        assert cursor.lastrowid is not None
        block_id = int(cursor.lastrowid)
        edge = store.connection.execute(
            """
            INSERT INTO edges(
                src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES ('fblock', ?, 1, 1, 3, 3, 2, 2, 'spill')
            """,
            (block_id,),
        )
        assert edge.lastrowid is not None
        store.edge_store.insert(int(edge.lastrowid), 1, Rect(3, 3, 2, 2))

        context = store._formula_reference_context(metadata)
        store._replace_circular_diagnostics(metadata, context)

        assert tuple(
            map(
                tuple,
                store.connection.execute(
                    "SELECT code, row, col, ref FROM diagnostics WHERE code = 'E_CIRCULAR'"
                ),
            )
        ) == (("E_CIRCULAR", 3, 2, "cell:Spill!B3"),)


def test_circular_adapter_reanalyzes_composite_range_hulls(tmp_path: Path) -> None:
    descriptor = SheetDescriptor(
        order=0,
        name="Model",
        sheet_id=1,
        rel_id="rId1",
        xml_part="xl/worksheets/sheet1.xml",
        kind="worksheet",
    )
    fixed_band = DefinedName(
        name="FixedBand",
        refers_to="Model!$B$2:$B$5",
        scope_sheet_order=None,
        kind="range",
        is_builtin=False,
        areas=(NameArea("Model", Rect(2, 5, 2, 2)),),
    )
    metadata = WorkbookMetadata(
        path=str(tmp_path / "composite.xlsx"),
        date1904=False,
        sheets=(descriptor,),
        defined_names=(fixed_band,),
    )
    with IndexStore(tmp_path / "composite.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))
        store.replace_defined_names(metadata)
        for row in range(3, 11):
            store.connection.execute(
                """
                INSERT INTO cells(
                    sheet_id, row, col, ref, value, value_type, formula,
                    style_idx, formula_kind, shared_index, array_ref, data_table
                ) VALUES (1, ?, 5, ?, 0, 'number', ?, 0,
                          'normal', NULL, NULL, NULL)
                """,
                (row, f"E{row}", f"=SUM(A{row}:FixedBand)"),
            )
        for row in range(6, 11):
            store.connection.execute(
                """
                INSERT INTO cells(
                    sheet_id, row, col, ref, value, value_type, formula,
                    style_idx, formula_kind, shared_index, array_ref, data_table
                ) VALUES (1, ?, 1, ?, 0, 'number', ?, 0,
                          'normal', NULL, NULL, NULL)
                """,
                (row, f"A{row}", f"=E{row}"),
            )
        first = store.connection.execute(
            """
            INSERT INTO fblocks(
                sheet_id, n, r1c1, row_min, row_max, col_min, col_max,
                volatile, opaque
            ) VALUES (1, 0, '=SUM(RC[-4]:FixedBand)', 3, 10, 5, 5, 0, 0)
            """
        )
        second = store.connection.execute(
            """
            INSERT INTO fblocks(
                sheet_id, n, r1c1, row_min, row_max, col_min, col_max,
                volatile, opaque
            ) VALUES (1, 1, '=RC[4]', 6, 10, 1, 1, 0, 0)
            """
        )
        assert first.lastrowid is not None and second.lastrowid is not None
        edge_rows = (
            (int(first.lastrowid), 2, 10, 1, 2, "ref"),
            (int(second.lastrowid), 6, 10, 5, 5, "ref"),
        )
        for block_id, row_min, row_max, col_min, col_max, via in edge_rows:
            cursor = store.connection.execute(
                """
                INSERT INTO edges(
                    src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES ('fblock', ?, 1, 1, ?, ?, ?, ?, ?)
                """,
                (block_id, row_min, row_max, col_min, col_max, via),
            )
            assert cursor.lastrowid is not None
            store.edge_store.insert(
                int(cursor.lastrowid),
                1,
                Rect(row_min, row_max, col_min, col_max),
            )

        context = store._formula_reference_context(metadata)
        store._replace_circular_diagnostics(metadata, context)

        rows = store.connection.execute(
            "SELECT code, ref FROM diagnostics WHERE code = 'E_CIRCULAR'"
        ).fetchall()
        assert len(rows) == 1
        assert tuple(rows[0]) == ("E_CIRCULAR", "cell:Model!A6")


def _trace_symbols(node: TraceNode) -> set[str]:
    return {node.target.label}.union(*(_trace_symbols(child) for child in node.children))


def _trace_snapshot(result: TraceResult) -> dict[str, object]:
    return {
        "direction": result.direction,
        "nodeCount": result.node_count,
        "edgeCount": result.edge_count,
        "truncated": result.truncated,
        "root": _trace_node_snapshot(result.root),
    }


def _trace_node_snapshot(node: TraceNode) -> dict[str, object]:
    return {
        "kind": node.target.kind,
        "symbol": node.target.symbol,
        "ref": node.target.ref,
        "via": node.via,
        "childCount": node.child_count,
        "children": [_trace_node_snapshot(child) for child in node.children],
    }


def _path_snapshot(result: PathResult) -> dict[str, object]:
    return {
        "connected": result.connected,
        "truncated": result.truncated,
        "paths": [
            [{"symbol": step.symbol, "via": step.via} for step in path] for path in result.paths
        ],
    }
