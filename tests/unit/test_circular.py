from __future__ import annotations

import random
from collections import Counter
from collections.abc import Callable, Iterable

import pytest

from excel_lsp.core.graph.circular import (
    BlockKey,
    CellDependency,
    CellNode,
    CircularBlock,
    CircularEdge,
    CircularLimits,
    _RectangleIndex,
    detect_circular_references,
    iterative_tarjan_scc,
)
from excel_lsp.core.models import Rect


def _block(
    sheet: int,
    ordinal: int,
    row_min: int,
    row_max: int,
    col_min: int,
    col_max: int,
) -> CircularBlock:
    return CircularBlock(BlockKey(sheet, ordinal), Rect(row_min, row_max, col_min, col_max))


def _edge(source: CircularBlock, destination: CircularBlock, *, via: str = "ref") -> CircularEdge:
    return CircularEdge(
        source=source.key,
        dst_sheet_id=destination.key.sheet_id,
        dst_rect=destination.rect,
        via=via,
    )


def _dependency(cell: CellNode) -> CellDependency:
    return CellDependency(cell.sheet_id, Rect(cell.row, cell.row, cell.col, cell.col))


def _mapping_resolver(
    mapping: dict[CellNode, tuple[CellDependency, ...]],
):
    def resolve(cell: CellNode) -> Iterable[CellDependency]:
        return mapping.get(cell, ())

    return resolve


def test_iterative_tarjan_handles_more_than_one_thousand_nodes_without_recursion() -> None:
    nodes = tuple(BlockKey(0, index) for index in range(1_505))
    graph = {node: (nodes[(index + 1) % len(nodes)],) for index, node in enumerate(nodes)}

    assert iterative_tarjan_scc(graph) == (nodes,)


def test_iterative_tarjan_matches_brute_force_random_graphs() -> None:
    generator = random.Random(0xC1AC)
    for size in range(1, 11):
        nodes = tuple(BlockKey(0, index) for index in range(size))
        for _case in range(30):
            graph = {
                node: tuple(target for target in nodes if generator.random() < 0.22)
                for node in nodes
            }
            expected = _brute_force_sccs(graph)
            actual = iterative_tarjan_scc(dict(reversed(tuple(graph.items()))))
            assert actual == expected


def test_rectangle_index_prunes_thousands_of_unrelated_blocks() -> None:
    blocks = tuple(_block(0, index, 1, 1, index + 1, index + 1) for index in range(6_000))
    target = blocks[4_731]
    inspected: list[CircularBlock] = []

    matches = _RectangleIndex(blocks).query(
        0,
        target.rect,
        inspect=inspected.append,
    )

    assert matches == (target,)
    assert len(inspected) <= 8


def test_coarse_and_exact_queries_both_prune_thousands_of_decoys(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    decoys = tuple(_block(0, index, 1, 1, index + 1, index + 1) for index in range(6_000))
    first = _block(0, 6_000, 1, 1, 10_000, 10_000)
    second = _block(0, 6_001, 1, 1, 10_001, 10_001)
    first_cell = CellNode(0, 1, 10_000)
    second_cell = CellNode(0, 1, 10_001)
    mapping = {
        first_cell: (_dependency(second_cell),),
        second_cell: (_dependency(first_cell),),
    }
    candidate_inspections: list[int] = []
    original_query = _RectangleIndex.query

    def instrumented_query(
        self: _RectangleIndex,
        sheet_id: int,
        rect: Rect,
        *,
        inspect: Callable[[CircularBlock], None] | None = None,
    ) -> tuple[CircularBlock, ...]:
        inspected: list[CircularBlock] = []

        def record(block: CircularBlock) -> None:
            inspected.append(block)
            if inspect is not None:
                inspect(block)

        result = original_query(self, sheet_id, rect, inspect=record)
        candidate_inspections.append(len(inspected))
        return result

    monkeypatch.setattr(_RectangleIndex, "query", instrumented_query)

    diagnostics = detect_circular_references(
        (*decoys, first, second),
        (_edge(first, second), _edge(second, first)),
        _mapping_resolver(mapping),
    )

    assert diagnostics[0].code == "E_CIRCULAR"
    assert len(candidate_inspections) == 4
    assert max(candidate_inspections[:2]) <= 8  # Coarse queries use the global index.
    assert candidate_inspections[2:] == [2, 2]  # Exact queries use the SCC-local index.


def test_stage2a_reports_one_canonical_error_for_self_inclusion() -> None:
    block = _block(0, 0, 1, 2, 1, 1)
    edges = (_edge(block, block),)
    mapping = {
        CellNode(0, 1, 1): (_dependency(CellNode(0, 1, 1)),),
        CellNode(0, 2, 1): (_dependency(CellNode(0, 2, 1)),),
    }

    diagnostics = detect_circular_references((block,), edges, _mapping_resolver(mapping))

    assert len(diagnostics) == 1
    diagnostic = diagnostics[0]
    assert (diagnostic.severity, diagnostic.code) == ("error", "E_CIRCULAR")
    assert diagnostic.ref == CellNode(0, 1, 1)
    assert diagnostic.related == (CellNode(0, 1, 1), CellNode(0, 1, 1))
    assert diagnostic.candidate_blocks == (block.key,)


def test_stage2b_proves_a_two_block_cycle_with_a_deterministic_path() -> None:
    left = _block(0, 0, 1, 1, 1, 1)
    right = _block(0, 1, 1, 1, 2, 2)
    left_cell = CellNode(0, 1, 1)
    right_cell = CellNode(0, 1, 2)
    mapping = {
        left_cell: (_dependency(right_cell),),
        right_cell: (_dependency(left_cell),),
    }

    diagnostics = detect_circular_references(
        (right, left),
        (_edge(right, left), _edge(left, right)),
        _mapping_resolver(mapping),
    )

    assert len(diagnostics) == 1
    assert diagnostics[0].code == "E_CIRCULAR"
    assert diagnostics[0].related == (left_cell, right_cell, left_cell)
    assert diagnostics[0].candidate_blocks == (left.key, right.key)


def test_stage2b_proves_a_cross_cell_cycle_inside_one_formula_block() -> None:
    block = _block(0, 0, 1, 2, 1, 1)
    first = CellNode(0, 1, 1)
    second = CellNode(0, 2, 1)
    mapping = {
        first: (_dependency(second),),
        second: (_dependency(first),),
    }

    diagnostics = detect_circular_references(
        (block,),
        (_edge(block, block),),
        _mapping_resolver(mapping),
    )

    assert len(diagnostics) == 1
    assert diagnostics[0].code == "E_CIRCULAR"
    assert diagnostics[0].related == (first, second, first)


def test_stage2b_proves_a_cross_sheet_cycle() -> None:
    first = _block(4, 0, 7, 7, 3, 3)
    second = _block(9, 0, 2, 2, 8, 8)
    first_cell = CellNode(4, 7, 3)
    second_cell = CellNode(9, 2, 8)
    mapping = {
        first_cell: (_dependency(second_cell),),
        second_cell: (_dependency(first_cell),),
    }

    diagnostics = detect_circular_references(
        (second, first),
        (_edge(first, second), _edge(second, first)),
        _mapping_resolver(mapping),
    )

    assert diagnostics[0].related == (first_cell, second_cell, first_cell)


def test_running_total_coarse_self_overlap_returns_nothing_and_never_runs_stage2b() -> None:
    running_total = _block(0, 0, 2, 1_501, 2, 2)
    coarse_overlap = CircularEdge(
        running_total.key,
        0,
        Rect(2, 1_500, 2, 2),
    )
    calls: Counter[CellNode] = Counter()

    def resolve(cell: CellNode) -> Iterable[CellDependency]:
        calls[cell] += 1
        if cell.row == 2:
            return ()
        return (CellDependency(0, Rect(2, cell.row - 1, 2, 2)),)

    diagnostics = detect_circular_references((running_total,), (coarse_overlap,), resolve)

    assert diagnostics == ()
    assert len(calls) == 1_500
    assert set(calls.values()) == {1}


def test_multi_block_coarse_false_cycle_returns_nothing_when_exact_graph_is_acyclic() -> None:
    first = _block(0, 0, 1, 1, 1, 1)
    second = _block(0, 1, 1, 1, 2, 2)
    first_cell = CellNode(0, 1, 1)
    second_cell = CellNode(0, 1, 2)
    mapping = {first_cell: (_dependency(second_cell),)}

    diagnostics = detect_circular_references(
        (first, second),
        (_edge(first, second), _edge(second, first)),
        _mapping_resolver(mapping),
    )

    assert diagnostics == ()


def test_opaque_and_external_edges_are_excluded_before_exact_resolution() -> None:
    first = _block(0, 0, 1, 1, 1, 1)
    second = _block(0, 1, 1, 1, 2, 2)
    edges = (
        _edge(first, second, via="opaque:INDIRECT"),
        CircularEdge(second.key, None, None, "external:budget.xlsx"),
    )

    def should_not_run(_cell: CellNode) -> Iterable[CellDependency]:
        raise AssertionError("non-concrete edges must not create a candidate SCC")

    assert detect_circular_references((first, second), edges, should_not_run) == ()


def test_stage2b_uses_at_most_64_deterministic_corner_and_sample_seeds() -> None:
    first = _block(0, 0, 1, 10, 1, 10)
    second = _block(0, 1, 20, 29, 1, 10)

    def repeated_calls(
        block_order: tuple[CircularBlock, ...], edge_order: tuple[CircularEdge, ...]
    ) -> tuple[CellNode, ...]:
        calls: Counter[CellNode] = Counter()

        def resolve(cell: CellNode) -> Iterable[CellDependency]:
            calls[cell] += 1
            return ()

        diagnostics = detect_circular_references(block_order, edge_order, resolve)
        assert len(diagnostics) == 1
        assert diagnostics[0].code == "W_POSSIBLE_CIRCULAR"
        return tuple(sorted(cell for cell, count in calls.items() if count == 2))

    forward_edges = (_edge(first, second), _edge(second, first))
    forward = repeated_calls((first, second), forward_edges)
    reversed_result = repeated_calls((second, first), tuple(reversed(forward_edges)))

    assert len(forward) == 64
    assert forward == reversed_result
    assert CellNode(0, 1, 1) in forward
    assert CellNode(0, 29, 10) in forward


def test_hidden_unsampled_two_block_cycle_downgrades_to_possible_circular() -> None:
    first = _block(0, 0, 1, 10, 1, 10)
    second = _block(0, 1, 20, 29, 1, 10)
    edges = (_edge(first, second), _edge(second, first))
    calls: Counter[CellNode] = Counter()

    def empty_resolver(cell: CellNode) -> Iterable[CellDependency]:
        calls[cell] += 1
        return ()

    baseline = detect_circular_references((first, second), edges, empty_resolver)
    assert baseline[0].code == "W_POSSIBLE_CIRCULAR"
    sampled = {cell for cell, count in calls.items() if count == 2}
    hidden_first = next(
        CellNode(0, row, col)
        for row in range(1, 11)
        for col in range(1, 11)
        if CellNode(0, row, col) not in sampled
    )
    hidden_second = next(
        CellNode(0, row, col)
        for row in range(20, 30)
        for col in range(1, 11)
        if CellNode(0, row, col) not in sampled
    )
    mapping = {
        hidden_first: (_dependency(hidden_second),),
        hidden_second: (_dependency(hidden_first),),
    }

    diagnostics = detect_circular_references(
        (first, second),
        edges,
        _mapping_resolver(mapping),
    )

    assert len(diagnostics) == 1
    assert diagnostics[0].code == "W_POSSIBLE_CIRCULAR"


@pytest.mark.parametrize(("maximum", "expected_code"), [(3, "W_POSSIBLE_CIRCULAR"), (4, None)])
def test_stage2b_charges_exactly_one_expansion_per_distinct_reachable_cell(
    maximum: int,
    expected_code: str | None,
) -> None:
    first = _block(0, 0, 1, 3, 1, 1)
    second = _block(0, 1, 1, 1, 2, 2)
    a1 = CellNode(0, 1, 1)
    a2 = CellNode(0, 2, 1)
    a3 = CellNode(0, 3, 1)
    b1 = CellNode(0, 1, 2)
    mapping = {
        a1: (_dependency(a2),),
        a2: (_dependency(a3),),
        a3: (_dependency(b1),),
    }

    diagnostics = detect_circular_references(
        (first, second),
        (_edge(first, second), _edge(second, first)),
        _mapping_resolver(mapping),
        limits=CircularLimits(max_seeds=1, max_expansions=maximum),
    )

    if expected_code is None:
        assert diagnostics == ()
    else:
        assert len(diagnostics) == 1
        assert diagnostics[0].code == expected_code
        assert diagnostics[0].severity == "warn"
        assert "verify in Excel" in diagnostics[0].message


def test_exact_cycle_wins_over_a_later_bound_exhaustion() -> None:
    first = _block(0, 0, 1, 3, 1, 1)
    second = _block(0, 1, 1, 3, 2, 2)
    a1 = CellNode(0, 1, 1)
    a2 = CellNode(0, 2, 1)
    mapping = {
        a1: (_dependency(a2),),
        a2: (_dependency(a1), CellDependency(0, second.rect)),
    }

    diagnostics = detect_circular_references(
        (first, second),
        (_edge(first, second), _edge(second, first)),
        _mapping_resolver(mapping),
        limits=CircularLimits(max_seeds=1, max_expansions=2),
    )

    assert len(diagnostics) == 1
    assert diagnostics[0].code == "E_CIRCULAR"
    assert diagnostics[0].related == (a1, a2, a1)


def test_invalid_limits_and_missing_edge_sources_fail_closed() -> None:
    with pytest.raises(ValueError, match="max_seeds"):
        CircularLimits(max_seeds=65)
    with pytest.raises(ValueError, match="max_expansions"):
        CircularLimits(max_expansions=100_001)

    block = _block(0, 0, 1, 1, 1, 1)
    edge = CircularEdge(BlockKey(0, 99), 0, block.rect)
    with pytest.raises(ValueError, match="no formula block"):
        detect_circular_references((block,), (edge,), lambda _cell: ())


def _brute_force_sccs(
    graph: dict[BlockKey, tuple[BlockKey, ...]],
) -> tuple[tuple[BlockKey, ...], ...]:
    reachability: dict[BlockKey, set[BlockKey]] = {}
    for start in graph:
        seen = {start}
        pending = [start]
        while pending:
            node = pending.pop()
            for target in graph[node]:
                if target not in seen:
                    seen.add(target)
                    pending.append(target)
        reachability[start] = seen

    remaining = set(graph)
    components: list[tuple[BlockKey, ...]] = []
    while remaining:
        first = min(remaining)
        component = tuple(
            node
            for node in sorted(remaining)
            if node in reachability[first] and first in reachability[node]
        )
        components.append(component)
        remaining.difference_update(component)
    return tuple(sorted(components))
