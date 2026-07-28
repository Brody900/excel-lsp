"""Pure, bounded circular-reference detection over coarse formula blocks.

Persistence adapters deliberately live outside this module.  Callers provide
small immutable block/edge records plus a resolver that translates one formula
cell into its exact dependency rectangles.

The historical phase text says a singleton coarse false overlap never needs
stage 2b.  That is safe only after an exact acyclicity proof: one formula block
can contain a cross-cell cycle.  This implementation conservatively uses a
strict coordinate-order proof for the running-total case and bounded stage 2b
for every other singleton.
"""

from __future__ import annotations

from collections import deque
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass
from typing import Generic, Literal, TypeVar

from excel_lsp.core.models import Rect

_MAX_SEEDS = 64
_MAX_EXPANSIONS = 100_000
_RECTANGLE_LEAF_SIZE = 8


@dataclass(frozen=True, order=True, slots=True)
class BlockKey:
    """Workbook-stable identity for one sheet-local formula block."""

    sheet_id: int
    ordinal: int

    def __post_init__(self) -> None:
        if type(self.sheet_id) is not int or self.sheet_id < 0:
            raise ValueError("block sheet id must be a nonnegative integer")
        if type(self.ordinal) is not int or self.ordinal < 0:
            raise ValueError("block ordinal must be a nonnegative integer")


@dataclass(frozen=True, order=True, slots=True)
class CellNode:
    """One exact one-based formula-cell coordinate."""

    sheet_id: int
    row: int
    col: int

    def __post_init__(self) -> None:
        if type(self.sheet_id) is not int or self.sheet_id < 0:
            raise ValueError("cell sheet id must be a nonnegative integer")
        if type(self.row) is not int or type(self.col) is not int:
            raise ValueError("cell coordinates must be integers")
        if not 1 <= self.row <= 1_048_576 or not 1 <= self.col <= 16_384:
            raise ValueError("cell coordinate exceeds Excel bounds")


@dataclass(frozen=True, slots=True)
class CircularBlock:
    """Coarse formula block consumed by circular detection."""

    key: BlockKey
    rect: Rect


@dataclass(frozen=True, slots=True)
class CircularEdge:
    """One coarse dependency edge; non-concrete edges are ignored."""

    source: BlockKey
    dst_sheet_id: int | None
    dst_rect: Rect | None
    via: str = "ref"

    def __post_init__(self) -> None:
        if (self.dst_sheet_id is None) != (self.dst_rect is None):
            raise ValueError("edge destination sheet and rectangle must coexist")
        if self.dst_sheet_id is not None and self.dst_sheet_id < 0:
            raise ValueError("edge destination sheet id must be nonnegative")
        if not self.via:
            raise ValueError("edge via must not be empty")


@dataclass(frozen=True, slots=True)
class CellDependency:
    """Exact dependency rectangle returned for one translated formula cell."""

    dst_sheet_id: int | None
    dst_rect: Rect | None
    via: str = "ref"

    def __post_init__(self) -> None:
        if (self.dst_sheet_id is None) != (self.dst_rect is None):
            raise ValueError("dependency destination sheet and rectangle must coexist")
        if self.dst_sheet_id is not None and self.dst_sheet_id < 0:
            raise ValueError("dependency destination sheet id must be nonnegative")
        if not self.via:
            raise ValueError("dependency via must not be empty")


@dataclass(frozen=True, slots=True)
class CircularLimits:
    """Frozen upper bounds, with smaller values available to deterministic tests."""

    max_seeds: int = _MAX_SEEDS
    max_expansions: int = _MAX_EXPANSIONS

    def __post_init__(self) -> None:
        if type(self.max_seeds) is not int or not 1 <= self.max_seeds <= _MAX_SEEDS:
            raise ValueError(f"max_seeds must be between 1 and {_MAX_SEEDS}")
        if type(self.max_expansions) is not int or not 1 <= self.max_expansions <= _MAX_EXPANSIONS:
            raise ValueError(f"max_expansions must be between 1 and {_MAX_EXPANSIONS}")


_DEFAULT_LIMITS = CircularLimits()


@dataclass(frozen=True, slots=True)
class CircularDiagnostic:
    """Canonical circular result for one coarse candidate SCC."""

    severity: Literal["error", "warn"]
    code: Literal["E_CIRCULAR", "W_POSSIBLE_CIRCULAR"]
    ref: CellNode
    related: tuple[CellNode, ...]
    candidate_blocks: tuple[BlockKey, ...]
    message: str

    def __post_init__(self) -> None:
        expected = "error" if self.code == "E_CIRCULAR" else "warn"
        if self.severity != expected:
            raise ValueError("circular diagnostic severity does not match its code")
        if not self.related or not self.candidate_blocks or not self.message:
            raise ValueError("circular diagnostic details must not be empty")


ExactDependencyResolver = Callable[[CellNode], Iterable[CellDependency]]

_NodeT = TypeVar("_NodeT")


@dataclass(slots=True)
class _TarjanFrame(Generic[_NodeT]):
    node: _NodeT
    neighbors: tuple[_NodeT, ...]
    offset: int = 0


@dataclass(frozen=True, slots=True)
class _ExactExpansion:
    adjacency: Mapping[CellNode, tuple[CellNode, ...]]
    exhausted: bool
    complete: bool


@dataclass(frozen=True, slots=True)
class _SelfCheck:
    hit: CellNode | None
    all_internal_before: bool
    all_internal_after: bool


@dataclass(frozen=True, slots=True)
class _RectangleNode:
    bounds: Rect
    blocks: tuple[CircularBlock, ...] = ()
    left: _RectangleNode | None = None
    right: _RectangleNode | None = None


class _RectangleIndex:
    """Deterministic per-sheet static rectangle index for formula blocks."""

    def __init__(self, blocks: Sequence[CircularBlock]) -> None:
        grouped: dict[int, list[CircularBlock]] = {}
        for block in blocks:
            grouped.setdefault(block.key.sheet_id, []).append(block)
        self._roots: dict[int, _RectangleNode] = {
            sheet_id: _build_rectangle_tree(tuple(sheet_blocks))
            for sheet_id, sheet_blocks in grouped.items()
        }

    def query(
        self,
        sheet_id: int,
        rect: Rect,
        *,
        inspect: Callable[[CircularBlock], None] | None = None,
    ) -> tuple[CircularBlock, ...]:
        """Return actual overlaps, pruning unrelated spatial subtrees."""
        root = self._roots.get(sheet_id)
        if root is None:
            return ()
        matches: list[CircularBlock] = []
        pending = [root]
        while pending:
            node = pending.pop()
            if not node.bounds.intersects(rect):
                continue
            if node.blocks:
                for block in node.blocks:
                    if inspect is not None:
                        inspect(block)
                    if block.rect.intersects(rect):
                        matches.append(block)
                continue
            if node.right is not None:
                pending.append(node.right)
            if node.left is not None:
                pending.append(node.left)
        return tuple(sorted(matches, key=_block_sort_key))


def iterative_tarjan_scc(
    graph: Mapping[BlockKey, Iterable[BlockKey]],
) -> tuple[tuple[BlockKey, ...], ...]:
    """Return deterministic SCCs without recursion, including isolated nodes."""
    adjacency: dict[BlockKey, tuple[BlockKey, ...]] = {}
    all_nodes = set(graph)
    for node, neighbors in graph.items():
        materialized = tuple(sorted(set(neighbors)))
        adjacency[node] = materialized
        all_nodes.update(materialized)
    for node in all_nodes:
        adjacency.setdefault(node, ())
    return _tarjan(adjacency, key=lambda item: (item.sheet_id, item.ordinal))


def detect_circular_references(
    blocks: Sequence[CircularBlock],
    edges: Sequence[CircularEdge],
    resolve_exact: ExactDependencyResolver,
    *,
    limits: CircularLimits = _DEFAULT_LIMITS,
) -> tuple[CircularDiagnostic, ...]:
    """Detect exact cell cycles behind coarse block-graph candidates.

    Stage 2a streams every cell of each candidate SCC and proves direct
    self-inclusion.  Singleton blocks stop early only when all exact internal
    edges share a strict coordinate direction; all other candidates continue
    through bounded exact-cell expansion.
    """
    ordered_blocks, block_by_key = _validated_blocks(blocks)
    rectangle_index = _RectangleIndex(ordered_blocks)
    adjacency = _coarse_adjacency(block_by_key, edges, rectangle_index)
    candidate_sccs = tuple(
        component
        for component in iterative_tarjan_scc(adjacency)
        if _is_candidate_component(component, adjacency)
    )
    results: list[CircularDiagnostic] = []
    for component in candidate_sccs:
        component_blocks = tuple(block_by_key[key] for key in component)
        self_check = _check_self_inclusion(component_blocks, resolve_exact)
        if self_check.hit is not None:
            results.append(_error_diagnostic(component, (self_check.hit, self_check.hit)))
            continue
        if len(component) == 1 and (
            self_check.all_internal_before or self_check.all_internal_after
        ):
            # A uniform strict coordinate ordering is a complete acyclicity
            # proof for the singleton's exact internal dependencies.  This is
            # the streaming running-total guard: every precedent lies before
            # its formula cell, so no exact cell graph needs to be allocated.
            continue

        seeds = _deterministic_seeds(component_blocks, limits.max_seeds)
        component_rectangle_index = _RectangleIndex(component_blocks)
        expansion = _expand_exact_graph(
            component_blocks,
            seeds,
            resolve_exact,
            limits.max_expansions,
            component_rectangle_index,
        )
        cycle = _find_canonical_cycle(expansion.adjacency)
        if cycle is not None:
            results.append(_error_diagnostic(component, cycle))
        elif expansion.exhausted or not expansion.complete:
            results.append(
                CircularDiagnostic(
                    severity="warn",
                    code="W_POSSIBLE_CIRCULAR",
                    ref=seeds[0],
                    related=seeds,
                    candidate_blocks=component,
                    message="Possible circular reference; verify in Excel.",
                )
            )
    return tuple(sorted(results, key=_diagnostic_sort_key))


def _validated_blocks(
    blocks: Sequence[CircularBlock],
) -> tuple[tuple[CircularBlock, ...], dict[BlockKey, CircularBlock]]:
    ordered = tuple(sorted(blocks, key=_block_sort_key))
    by_key: dict[BlockKey, CircularBlock] = {}
    for block in ordered:
        if block.key in by_key:
            raise ValueError(f"duplicate circular block key: {block.key!r}")
        by_key[block.key] = block
    return ordered, by_key


def _coarse_adjacency(
    block_by_key: Mapping[BlockKey, CircularBlock],
    edges: Sequence[CircularEdge],
    rectangle_index: _RectangleIndex,
) -> dict[BlockKey, tuple[BlockKey, ...]]:
    mutable: dict[BlockKey, set[BlockKey]] = {key: set() for key in block_by_key}
    for edge in sorted(edges, key=_edge_sort_key):
        if edge.source not in block_by_key:
            raise ValueError(f"edge source has no formula block: {edge.source!r}")
        if not _is_concrete(edge.dst_sheet_id, edge.dst_rect, edge.via):
            continue
        assert edge.dst_sheet_id is not None
        assert edge.dst_rect is not None
        for target in rectangle_index.query(edge.dst_sheet_id, edge.dst_rect):
            mutable[edge.source].add(target.key)
    return {key: tuple(sorted(targets)) for key, targets in sorted(mutable.items())}


def _check_self_inclusion(
    blocks: Sequence[CircularBlock],
    resolve_exact: ExactDependencyResolver,
) -> _SelfCheck:
    track_order_proof = len(blocks) == 1
    all_before = True
    all_after = True
    for block in blocks:
        for cell in _block_cells(block):
            for dependency in resolve_exact(cell):
                if _dependency_contains(dependency, cell):
                    return _SelfCheck(cell, False, False)
                if not track_order_proof:
                    continue
                bounds = _internal_dependency_bounds(dependency, block)
                if bounds is None:
                    continue
                minimum, maximum = bounds
                if maximum >= cell:
                    all_before = False
                if minimum <= cell:
                    all_after = False
    return _SelfCheck(None, all_before, all_after)


def _deterministic_seeds(
    blocks: Sequence[CircularBlock],
    maximum: int,
) -> tuple[CellNode, ...]:
    total_cells = sum(_rect_area(block.rect) for block in blocks)
    if total_cells <= maximum:
        return tuple(cell for block in blocks for cell in _block_cells(block))

    corners = sorted(
        {
            CellNode(block.key.sheet_id, row, col)
            for block in blocks
            for row in (block.rect.row_min, block.rect.row_max)
            for col in (block.rect.col_min, block.rect.col_max)
        }
    )
    corner_budget = min(len(corners), maximum // 2)
    selected = _evenly_select(corners, corner_budget)
    selected_set = set(selected)
    sample_count = maximum - len(selected)
    samples = tuple(
        _cell_at_ordinal(blocks, ordinal)
        for ordinal in _evenly_spaced_ordinals(total_cells, sample_count)
    )
    for sample in samples:
        if sample not in selected_set:
            selected.append(sample)
            selected_set.add(sample)

    # Quantile collisions with corners can leave capacity.  Deterministic
    # forward sampling fills it without ever materializing every SCC cell.
    if len(selected) < maximum:
        for ordinal in _evenly_spaced_ordinals(total_cells, maximum * 2):
            sample = _cell_at_ordinal(blocks, ordinal)
            if sample not in selected_set:
                selected.append(sample)
                selected_set.add(sample)
                if len(selected) == maximum:
                    break
    return tuple(sorted(selected))


def _expand_exact_graph(
    blocks: Sequence[CircularBlock],
    seeds: tuple[CellNode, ...],
    resolve_exact: ExactDependencyResolver,
    maximum: int,
    rectangle_index: _RectangleIndex,
) -> _ExactExpansion:
    queue = deque(seeds)
    discovered = set(seeds)
    adjacency: dict[CellNode, tuple[CellNode, ...]] = {}
    charged_edges = 0
    if len(discovered) > maximum:
        return _ExactExpansion(adjacency, exhausted=True, complete=False)

    while queue:
        source = queue.popleft()
        targets: set[CellNode] = set()
        dependencies = sorted(resolve_exact(source), key=_dependency_sort_key)
        for dependency in dependencies:
            if not _is_concrete(
                dependency.dst_sheet_id,
                dependency.dst_rect,
                dependency.via,
            ):
                continue
            for target in _formula_cells_in_dependency(
                dependency,
                rectangle_index,
            ):
                if target in targets:
                    continue
                if charged_edges == maximum:
                    adjacency[source] = tuple(sorted(targets))
                    return _ExactExpansion(adjacency, exhausted=True, complete=False)
                charged_edges += 1
                targets.add(target)
                if target in discovered:
                    continue
                if len(discovered) == maximum:
                    adjacency[source] = tuple(sorted(targets))
                    return _ExactExpansion(adjacency, exhausted=True, complete=False)
                discovered.add(target)
                queue.append(target)
        adjacency[source] = tuple(sorted(targets))
    for cell in discovered:
        adjacency.setdefault(cell, ())
    expected_cells = sum(_rect_area(block.rect) for block in blocks)
    return _ExactExpansion(
        adjacency,
        exhausted=False,
        complete=len(discovered) == expected_cells,
    )


def _formula_cells_in_dependency(
    dependency: CellDependency,
    rectangle_index: _RectangleIndex,
) -> Iterable[CellNode]:
    assert dependency.dst_sheet_id is not None
    assert dependency.dst_rect is not None
    for block in rectangle_index.query(dependency.dst_sheet_id, dependency.dst_rect):
        intersection = _intersection(block.rect, dependency.dst_rect)
        if intersection is None:
            continue
        for row in range(intersection.row_min, intersection.row_max + 1):
            for col in range(intersection.col_min, intersection.col_max + 1):
                yield CellNode(block.key.sheet_id, row, col)


def _find_canonical_cycle(
    adjacency: Mapping[CellNode, tuple[CellNode, ...]],
) -> tuple[CellNode, ...] | None:
    cyclic = [
        component
        for component in _tarjan(
            adjacency,
            key=lambda item: (item.sheet_id, item.row, item.col),
        )
        if len(component) > 1 or (component and component[0] in adjacency.get(component[0], ()))
    ]
    if not cyclic:
        return None
    component = min(cyclic)
    allowed = set(component)
    start = component[0]
    if start in adjacency.get(start, ()):
        return start, start

    queue: deque[CellNode] = deque()
    parent: dict[CellNode, CellNode] = {}
    for neighbor in adjacency.get(start, ()):
        if neighbor in allowed and neighbor != start and neighbor not in parent:
            parent[neighbor] = start
            queue.append(neighbor)
    while queue:
        node = queue.popleft()
        for neighbor in adjacency.get(node, ()):
            if neighbor == start:
                return _reconstruct_cycle(start, node, parent)
            if neighbor in allowed and neighbor not in parent:
                parent[neighbor] = node
                queue.append(neighbor)
    raise RuntimeError("strongly connected cell component has no cycle path")


def _reconstruct_cycle(
    start: CellNode,
    end: CellNode,
    parent: Mapping[CellNode, CellNode],
) -> tuple[CellNode, ...]:
    reverse_path = [end]
    while reverse_path[-1] != start:
        reverse_path.append(parent[reverse_path[-1]])
    reverse_path.reverse()
    reverse_path.append(start)
    return tuple(reverse_path)


def _error_diagnostic(
    component: tuple[BlockKey, ...],
    cycle: tuple[CellNode, ...],
) -> CircularDiagnostic:
    return CircularDiagnostic(
        severity="error",
        code="E_CIRCULAR",
        ref=cycle[0],
        related=cycle,
        candidate_blocks=component,
        message="Circular reference detected.",
    )


def _tarjan(
    adjacency: Mapping[_NodeT, Iterable[_NodeT]],
    *,
    key: Callable[[_NodeT], tuple[int, ...]],
) -> tuple[tuple[_NodeT, ...], ...]:
    normalized: dict[_NodeT, tuple[_NodeT, ...]] = {
        node: tuple(sorted(set(targets), key=key)) for node, targets in adjacency.items()
    }
    all_nodes = set(normalized)
    for targets in normalized.values():
        all_nodes.update(targets)
    for node in all_nodes:
        normalized.setdefault(node, ())

    index: dict[_NodeT, int] = {}
    lowlink: dict[_NodeT, int] = {}
    active: list[_NodeT] = []
    on_active: set[_NodeT] = set()
    components: list[tuple[_NodeT, ...]] = []
    next_index = 0

    for root in sorted(all_nodes, key=key):
        if root in index:
            continue
        index[root] = next_index
        lowlink[root] = next_index
        next_index += 1
        active.append(root)
        on_active.add(root)
        frames = [_TarjanFrame(root, normalized[root])]

        while frames:
            frame = frames[-1]
            if frame.offset < len(frame.neighbors):
                neighbor = frame.neighbors[frame.offset]
                frame.offset += 1
                if neighbor not in index:
                    index[neighbor] = next_index
                    lowlink[neighbor] = next_index
                    next_index += 1
                    active.append(neighbor)
                    on_active.add(neighbor)
                    frames.append(_TarjanFrame(neighbor, normalized[neighbor]))
                elif neighbor in on_active:
                    lowlink[frame.node] = min(lowlink[frame.node], index[neighbor])
                continue

            frames.pop()
            if frames:
                parent = frames[-1].node
                lowlink[parent] = min(lowlink[parent], lowlink[frame.node])
            if lowlink[frame.node] != index[frame.node]:
                continue
            component: list[_NodeT] = []
            while True:
                member = active.pop()
                on_active.remove(member)
                component.append(member)
                if member == frame.node:
                    break
            components.append(tuple(sorted(component, key=key)))

    return tuple(sorted(components, key=lambda component: tuple(key(item) for item in component)))


def _is_candidate_component(
    component: tuple[BlockKey, ...],
    adjacency: Mapping[BlockKey, tuple[BlockKey, ...]],
) -> bool:
    if len(component) > 1:
        return True
    if not component:
        return False
    node = component[0]
    return node in adjacency[node]


def _dependency_contains(dependency: CellDependency, cell: CellNode) -> bool:
    if not _is_concrete(
        dependency.dst_sheet_id,
        dependency.dst_rect,
        dependency.via,
    ):
        return False
    assert dependency.dst_rect is not None
    return (
        dependency.dst_sheet_id == cell.sheet_id
        and dependency.dst_rect.row_min <= cell.row <= dependency.dst_rect.row_max
        and dependency.dst_rect.col_min <= cell.col <= dependency.dst_rect.col_max
    )


def _internal_dependency_bounds(
    dependency: CellDependency,
    block: CircularBlock,
) -> tuple[CellNode, CellNode] | None:
    if (
        not _is_concrete(
            dependency.dst_sheet_id,
            dependency.dst_rect,
            dependency.via,
        )
        or dependency.dst_sheet_id != block.key.sheet_id
    ):
        return None
    assert dependency.dst_rect is not None
    intersection = _intersection(dependency.dst_rect, block.rect)
    if intersection is None:
        return None
    return (
        CellNode(block.key.sheet_id, intersection.row_min, intersection.col_min),
        CellNode(block.key.sheet_id, intersection.row_max, intersection.col_max),
    )


def _is_concrete(sheet_id: int | None, rect: Rect | None, via: str) -> bool:
    if sheet_id is None or rect is None:
        return False
    normalized = via.casefold()
    return not normalized.startswith(("opaque", "external"))


def _block_cells(block: CircularBlock) -> Iterable[CellNode]:
    for row in range(block.rect.row_min, block.rect.row_max + 1):
        for col in range(block.rect.col_min, block.rect.col_max + 1):
            yield CellNode(block.key.sheet_id, row, col)


def _cell_at_ordinal(blocks: Sequence[CircularBlock], ordinal: int) -> CellNode:
    remaining = ordinal
    for block in blocks:
        area = _rect_area(block.rect)
        if remaining >= area:
            remaining -= area
            continue
        width = block.rect.col_max - block.rect.col_min + 1
        row_offset, col_offset = divmod(remaining, width)
        return CellNode(
            block.key.sheet_id,
            block.rect.row_min + row_offset,
            block.rect.col_min + col_offset,
        )
    raise IndexError("cell ordinal exceeds formula-block area")


def _evenly_select(items: Sequence[_NodeT], count: int) -> list[_NodeT]:
    if count == 0:
        return []
    return [items[index] for index in _evenly_spaced_ordinals(len(items), count)]


def _evenly_spaced_ordinals(total: int, count: int) -> tuple[int, ...]:
    if count <= 0 or total <= 0:
        return ()
    if count == 1:
        return (0,)
    count = min(count, total)
    return tuple(index * (total - 1) // (count - 1) for index in range(count))


def _intersection(left: Rect, right: Rect) -> Rect | None:
    if not left.intersects(right):
        return None
    return Rect(
        max(left.row_min, right.row_min),
        min(left.row_max, right.row_max),
        max(left.col_min, right.col_min),
        min(left.col_max, right.col_max),
    )


def _rect_area(rect: Rect) -> int:
    return (rect.row_max - rect.row_min + 1) * (rect.col_max - rect.col_min + 1)


def _build_rectangle_tree(blocks: tuple[CircularBlock, ...]) -> _RectangleNode:
    if not blocks:
        raise ValueError("rectangle-tree leaves must not be empty")
    bounds = Rect(
        min(block.rect.row_min for block in blocks),
        max(block.rect.row_max for block in blocks),
        min(block.rect.col_min for block in blocks),
        max(block.rect.col_max for block in blocks),
    )
    if len(blocks) <= _RECTANGLE_LEAF_SIZE:
        return _RectangleNode(bounds, tuple(sorted(blocks, key=_block_sort_key)))

    row_centers = tuple(block.rect.row_min + block.rect.row_max for block in blocks)
    col_centers = tuple(block.rect.col_min + block.rect.col_max for block in blocks)
    split_on_row = max(row_centers) - min(row_centers) >= max(col_centers) - min(col_centers)

    def split_key(block: CircularBlock) -> tuple[int, ...]:
        row_center = block.rect.row_min + block.rect.row_max
        col_center = block.rect.col_min + block.rect.col_max
        if split_on_row:
            return row_center, col_center, *_block_sort_key(block)
        return col_center, row_center, *_block_sort_key(block)

    ordered = tuple(sorted(blocks, key=split_key))
    midpoint = len(ordered) // 2
    return _RectangleNode(
        bounds,
        left=_build_rectangle_tree(ordered[:midpoint]),
        right=_build_rectangle_tree(ordered[midpoint:]),
    )


def _block_sort_key(block: CircularBlock) -> tuple[int, int, int, int, int, int]:
    return (
        block.key.sheet_id,
        block.rect.row_min,
        block.rect.col_min,
        block.rect.row_max,
        block.rect.col_max,
        block.key.ordinal,
    )


def _edge_sort_key(edge: CircularEdge) -> tuple[object, ...]:
    rect = edge.dst_rect
    return (
        edge.source,
        -1 if edge.dst_sheet_id is None else edge.dst_sheet_id,
        -1 if rect is None else rect.row_min,
        -1 if rect is None else rect.col_min,
        -1 if rect is None else rect.row_max,
        -1 if rect is None else rect.col_max,
        edge.via,
    )


def _dependency_sort_key(dependency: CellDependency) -> tuple[object, ...]:
    rect = dependency.dst_rect
    return (
        -1 if dependency.dst_sheet_id is None else dependency.dst_sheet_id,
        -1 if rect is None else rect.row_min,
        -1 if rect is None else rect.col_min,
        -1 if rect is None else rect.row_max,
        -1 if rect is None else rect.col_max,
        dependency.via,
    )


def _diagnostic_sort_key(diagnostic: CircularDiagnostic) -> tuple[object, ...]:
    return diagnostic.candidate_blocks, diagnostic.ref, diagnostic.code, diagnostic.related


__all__ = [
    "BlockKey",
    "CellDependency",
    "CellNode",
    "CircularBlock",
    "CircularDiagnostic",
    "CircularEdge",
    "CircularLimits",
    "ExactDependencyResolver",
    "detect_circular_references",
    "iterative_tarjan_scc",
]
