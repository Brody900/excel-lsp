"""Sparse worksheet region detection, header inference, and column profiles.

The OOXML parser remains single-pass.  This module operates on a repeatable,
coordinate-ordered sparse cell stream after parser metadata (notably native
Excel Tables) is known.  No dense worksheet grid is constructed.
"""

from __future__ import annotations

import math
from bisect import bisect_left, bisect_right
from collections import Counter, defaultdict, deque
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime, time
from heapq import heappop, heappush
from typing import Literal, NoReturn, TypeAlias

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.models import CellScalar, CellValueType, Rect, SheetParseSummary, TableInfo
from excel_lsp.core.parse.coordinates import column_label, contains, parse_rect
from excel_lsp.core.parse.styles import StyleCatalog
from excel_lsp.core.symbols import (
    deduplicate_normalized_headers,
    normalize_header,
    sheet_symbol_id,
)

RegionKind: TypeAlias = Literal["table", "region"]
ColumnDType: TypeAlias = Literal["int", "float", "date", "str", "bool", "mixed", "empty"]
CellStreamFactory: TypeAlias = Callable[[], Iterable["RegionCell"]]
HeaderCellProvider: TypeAlias = Callable[[Sequence[Rect]], Iterable["RegionCell"]]
ColumnProfileProvider: TypeAlias = Callable[
    [Sequence["ColumnProfileRequest"], int, "RegionOptions"],
    tuple[tuple["ColumnProfile", ...], ...],
]

_HEADER_WEIGHTS = (0.30, 0.25, 0.20, 0.20, 0.05)
HEADER_BODY_PREVIEW = 24
_OVERLAP_BUCKET_WIDTH = 64
_OVERLAP_WIDE_BUCKETS = 4
_RECT_GRID_BUCKET_SIZE = 8
_RECT_GRID_MAX_CELLS = 64
_RECT_BALANCED_AXIS_RATIO = 4
_RECT_BALANCED_AXIS_MIN = 8
_MAX_GAP_TOL = 8


def _new_header_cells() -> dict[tuple[int, int], RegionCell]:
    return {}


def _new_header_preview() -> dict[int, list[RegionCell]]:
    return defaultdict(list)


def _new_dtype_atoms() -> set[str]:
    return set()


def _new_distinct_values() -> set[tuple[str, str, str]]:
    return set()


@dataclass(frozen=True, slots=True)
class RegionOptions:
    """Deterministic region-analysis settings."""

    gap_tol: int = 1
    header_threshold: float = 0.55
    dtype_sample_limit: int = 200
    distinct_cap: int = 1_000
    large_sheet_threshold: int = 2_000_000
    large_dtype_sample_limit: int = 50

    def __post_init__(self) -> None:
        if type(self.gap_tol) is not int:
            raise ValueError("gap_tol must be an integer")
        if self.gap_tol < 0:
            raise ValueError("gap_tol must be nonnegative")
        if self.gap_tol > _MAX_GAP_TOL:
            raise ValueError(f"gap_tol must not exceed {_MAX_GAP_TOL}")
        if not 0.0 <= self.header_threshold <= 1.0:
            raise ValueError("header_threshold must be between zero and one")
        for name in (
            "dtype_sample_limit",
            "distinct_cap",
            "large_sheet_threshold",
            "large_dtype_sample_limit",
        ):
            if getattr(self, name) < 1:
                raise ValueError(f"{name} must be positive")


@dataclass(frozen=True, slots=True)
class RegionCell:
    """The small cell projection consumed by region analysis."""

    row: int
    col: int
    value: CellScalar
    value_type: CellValueType
    style_idx: int = 0
    formula: str | None = None

    def __post_init__(self) -> None:
        if not 1 <= self.row <= 1_048_576:
            raise ValueError("cell row exceeds Excel bounds")
        if not 1 <= self.col <= 16_384:
            raise ValueError("cell column exceeds Excel bounds")
        if self.style_idx < 0:
            raise ValueError("style_idx must be nonnegative")


@dataclass(frozen=True, slots=True)
class ColumnProfile:
    """One persisted column description within a detected region."""

    idx: int
    header: str
    norm_header: str
    dtype: ColumnDType
    nonnull: int
    distinct_est: int


@dataclass(frozen=True, slots=True)
class ColumnProfileRequest:
    """One detected-region body whose column statistics must be profiled."""

    rect: Rect
    header_rows: int
    headers: tuple[str, ...]
    normalized_headers: tuple[str, ...]
    totals_rows: int = 0

    def __post_init__(self) -> None:
        width = _width(self.rect)
        if len(self.headers) != width or len(self.normalized_headers) != width:
            raise ValueError("column profile request headers must match region width")
        if not 0 <= self.header_rows <= _height(self.rect):
            raise ValueError("column profile request header rows exceed region height")
        if not 0 <= self.totals_rows <= _height(self.rect) - self.header_rows:
            raise ValueError("column profile request totals rows exceed region body")


@dataclass(frozen=True, slots=True)
class DetectedRegion:
    """One public region with its deterministic coordinate ordinal."""

    n: int
    rect: Rect
    header_rows: int
    kind: RegionKind
    list_object_name: str | None
    confidence: float
    columns: tuple[ColumnProfile, ...]


@dataclass(frozen=True, slots=True)
class RegionWarning:
    """A diagnostic emitted during structural analysis."""

    code: str
    ref: str
    message: str
    related: Mapping[str, int]
    severity: Literal["warn"] = "warn"


@dataclass(frozen=True, slots=True)
class RegionAnalysis:
    """Complete region result for one worksheet."""

    regions: tuple[DetectedRegion, ...]
    warnings: tuple[RegionWarning, ...] = ()


@dataclass(frozen=True, slots=True)
class _Run:
    row: int
    col_min: int
    col_max: int


class _RunSequence:
    """Immutable row-ordered runs with logarithmic slice bounds."""

    def __init__(self, runs: Sequence[_Run]) -> None:
        self.runs = tuple(runs)
        size = 1
        while size < len(self.runs):
            size *= 2
        self._size = size
        self._minimum_columns = [16_385] * (2 * size)
        self._maximum_columns = [0] * (2 * size)
        for offset, run in enumerate(self.runs):
            position = size + offset
            self._minimum_columns[position] = run.col_min
            self._maximum_columns[position] = run.col_max
        for position in range(size - 1, 0, -1):
            self._minimum_columns[position] = min(
                self._minimum_columns[position * 2],
                self._minimum_columns[position * 2 + 1],
            )
            self._maximum_columns[position] = max(
                self._maximum_columns[position * 2],
                self._maximum_columns[position * 2 + 1],
            )

    def lower_bound(self, start: int, stop: int, row: int) -> int:
        lower = start
        upper = stop
        while lower < upper:
            middle = (lower + upper) // 2
            if self.runs[middle].row < row:
                lower = middle + 1
            else:
                upper = middle
        return lower

    def upper_bound(self, start: int, stop: int, row: int) -> int:
        lower = start
        upper = stop
        while lower < upper:
            middle = (lower + upper) // 2
            if self.runs[middle].row <= row:
                lower = middle + 1
            else:
                upper = middle
        return lower

    def bounds(self, start: int, stop: int) -> Rect:
        if start >= stop:
            raise ValueError("run view must not be empty")
        left = start + self._size
        right = stop + self._size
        col_min = 16_385
        col_max = 0
        while left < right:
            if left & 1:
                col_min = min(col_min, self._minimum_columns[left])
                col_max = max(col_max, self._maximum_columns[left])
                left += 1
            if right & 1:
                right -= 1
                col_min = min(col_min, self._minimum_columns[right])
                col_max = max(col_max, self._maximum_columns[right])
            left //= 2
            right //= 2
        return Rect(
            self.runs[start].row,
            self.runs[stop - 1].row,
            col_min,
            col_max,
        )


@dataclass(frozen=True, slots=True)
class _RunView:
    source: _RunSequence
    start: int
    stop: int

    @classmethod
    def complete(cls, runs: Sequence[_Run]) -> _RunView:
        source = _RunSequence(runs)
        return cls(source, 0, len(source.runs))

    @property
    def bounds(self) -> Rect:
        return self.source.bounds(self.start, self.stop)

    def iter_runs(self) -> Iterable[_Run]:
        for index in range(self.start, self.stop):
            yield self.source.runs[index]

    def materialize(self) -> tuple[_Run, ...]:
        return self.source.runs[self.start : self.stop]


@dataclass(frozen=True, slots=True)
class _Span:
    row_min: int
    row_max: int
    col_min: int
    col_max: int

    @property
    def rect(self) -> Rect:
        return Rect(self.row_min, self.row_max, self.col_min, self.col_max)


@dataclass(frozen=True, slots=True)
class _SparseZone:
    runs: _RunView | None
    spans: tuple[_Span, ...]
    bounds: Rect

    @classmethod
    def create(
        cls,
        runs: _RunView | None,
        spans: Sequence[_Span],
    ) -> _SparseZone | None:
        ordered_spans = tuple(
            sorted(
                spans,
                key=lambda span: (
                    span.row_min,
                    span.col_min,
                    span.row_max,
                    span.col_max,
                ),
            )
        )
        bounds: Rect | None = None if runs is None else runs.bounds
        if ordered_spans:
            span_bounds = _rectangles_bounds(span.rect for span in ordered_spans)
            bounds = span_bounds if bounds is None else _bounding_rect(bounds, span_bounds)
        if bounds is None:
            return None
        return cls(runs, ordered_spans, bounds)


@dataclass(frozen=True, slots=True)
class _SparsePrimitive:
    member_id: int
    rect: Rect


@dataclass(frozen=True, slots=True)
class _SparseMemberBlock:
    member_ids: tuple[int, ...]
    index: _RectangleIndex


@dataclass(frozen=True, slots=True)
class _SparseComponent:
    runs: tuple[_Run, ...]
    spans: tuple[_Span, ...]
    bounds: Rect


@dataclass(frozen=True, slots=True)
class _Seed:
    rect: Rect
    kind: RegionKind
    table: TableInfo | None = None


@dataclass(frozen=True, slots=True)
class _HeaderDecision:
    rows: int
    confidence: float
    headers: tuple[str, ...]
    normalized_headers: tuple[str, ...]


@dataclass(slots=True)
class _HeaderEvidence:
    cells: dict[tuple[int, int], RegionCell] = field(default_factory=_new_header_cells)
    preview: dict[int, list[RegionCell]] = field(default_factory=_new_header_preview)


@dataclass(frozen=True, slots=True)
class _HeaderMergeSpan:
    col_min: int
    col_max: int
    anchor: tuple[int, int]


class _HeaderMergeView:
    """At-most-three-row merged-header lookup shared by all candidate depths."""

    def __init__(self, spans_by_row: Mapping[int, Sequence[_HeaderMergeSpan]]) -> None:
        self._spans_by_row = {
            row: tuple(sorted(spans, key=lambda span: (span.col_min, span.col_max, span.anchor)))
            for row, spans in spans_by_row.items()
        }
        self._starts_by_row = {
            row: tuple(span.col_min for span in spans) for row, spans in self._spans_by_row.items()
        }

    def anchor_at(self, row: int, col: int) -> tuple[int, int] | None:
        spans = self._spans_by_row.get(row)
        if spans is None:
            return None
        position = bisect_right(self._starts_by_row[row], col) - 1
        if position < 0 or spans[position].col_max < col:
            return None
        return spans[position].anchor


class _MergeHeaderIndex:
    """Workbook-sparse row, column, and direct-anchor lookup for merges."""

    def __init__(self, merges: Sequence[Rect]) -> None:
        self._merges = tuple(sorted(merges, key=_rect_key))
        self._rectangles = _RectangleIndex(self._merges)
        self._by_anchor = {(merge.row_min, merge.col_min): merge for merge in self._merges}

    def view(self, rect: Rect, header_rows: int) -> _HeaderMergeView:
        if header_rows < 1:
            return _HeaderMergeView({})
        window = Rect(
            rect.row_min,
            rect.row_min + header_rows - 1,
            rect.col_min,
            rect.col_max,
        )
        spans_by_row: dict[int, list[_HeaderMergeSpan]] = defaultdict(list)
        for merge_index in self._rectangles.intersections(window):
            merge = self._merges[merge_index]
            anchor = (merge.row_min, merge.col_min)
            canonical = self._by_anchor[anchor]
            for row in range(
                max(window.row_min, canonical.row_min),
                min(window.row_max, canonical.row_max) + 1,
            ):
                spans_by_row[row].append(
                    _HeaderMergeSpan(canonical.col_min, canonical.col_max, anchor)
                )
        return _HeaderMergeView(spans_by_row)


@dataclass(slots=True)
class _ColumnAccumulator:
    sample_limit: int
    sample_stride: int
    distinct_cap: int
    nonnull: int = 0
    eligible_seen: int = 0
    sampled: int = 0
    atoms: set[str] = field(default_factory=_new_dtype_atoms)
    distinct: set[tuple[str, str, str]] = field(default_factory=_new_distinct_values)

    def add(self, cell: RegionCell) -> None:
        if cell.value is None or cell.value_type == "blank":
            return
        self.nonnull += 1
        if len(self.distinct) < self.distinct_cap:
            self.distinct.add(_distinct_key(cell))
        sample_number = self.eligible_seen
        self.eligible_seen += 1
        if sample_number % self.sample_stride != 0 or self.sampled >= self.sample_limit:
            return
        self.sampled += 1
        self.atoms.add(_dtype_atom(cell))

    def dtype(self) -> ColumnDType:
        return column_dtype_from_atoms(self.atoms)


def column_dtype_from_atoms(atoms: Iterable[str]) -> ColumnDType:
    """Classify the frozen dtype atom set used by stream and SQL profilers."""
    values = set(atoms)
    if not values:
        return "empty"
    if values == {"int"}:
        return "int"
    if values <= {"int", "float"}:
        return "float"
    if values == {"date"}:
        return "date"
    if values == {"str"}:
        return "str"
    if values == {"bool"}:
        return "bool"
    return "mixed"


class _DisjointSet:
    def __init__(self, size: int) -> None:
        self.parent = list(range(size))
        self.rank = [0] * size

    def find(self, item: int) -> int:
        parent = self.parent[item]
        while parent != self.parent[parent]:
            parent = self.parent[parent]
        while item != parent:
            next_item = self.parent[item]
            self.parent[item] = parent
            item = next_item
        return parent

    def union(self, left: int, right: int) -> int:
        left_root = self.find(left)
        right_root = self.find(right)
        if left_root == right_root:
            return left_root
        if self.rank[left_root] < self.rank[right_root]:
            left_root, right_root = right_root, left_root
        self.parent[right_root] = left_root
        if self.rank[left_root] == self.rank[right_root]:
            self.rank[left_root] += 1
        return left_root


class _Fenwick:
    """Small dynamic prefix-count index over compressed coordinates."""

    def __init__(self, size: int) -> None:
        self._tree = [0] * (size + 1)

    def add(self, position: int, delta: int) -> None:
        index = position + 1
        while index < len(self._tree):
            self._tree[index] += delta
            index += index & -index

    def prefix(self, end: int) -> int:
        total = 0
        index = end
        while index:
            total += self._tree[index]
            index -= index & -index
        return total


class _IntervalNode:
    """One deterministic AVL node with a subtree maximum interval end."""

    __slots__ = ("end", "height", "item", "key", "left", "max_end", "right")

    def __init__(self, key: tuple[int, int], end: int, item: int) -> None:
        self.key = key
        self.end = end
        self.item = item
        self.height = 1
        self.max_end = end
        self.left: _IntervalNode | None = None
        self.right: _IntervalNode | None = None


def _interval_height(node: _IntervalNode | None) -> int:
    return 0 if node is None else node.height


def _interval_max_end(node: _IntervalNode | None) -> int:
    return 0 if node is None else node.max_end


def _refresh_interval_node(node: _IntervalNode) -> None:
    node.height = 1 + max(_interval_height(node.left), _interval_height(node.right))
    node.max_end = max(node.end, _interval_max_end(node.left), _interval_max_end(node.right))


def _rotate_interval_right(node: _IntervalNode) -> _IntervalNode:
    pivot = node.left
    if pivot is None:
        return node
    node.left = pivot.right
    pivot.right = node
    _refresh_interval_node(node)
    _refresh_interval_node(pivot)
    return pivot


def _rotate_interval_left(node: _IntervalNode) -> _IntervalNode:
    pivot = node.right
    if pivot is None:
        return node
    node.right = pivot.left
    pivot.left = node
    _refresh_interval_node(node)
    _refresh_interval_node(pivot)
    return pivot


def _rebalance_interval_node(node: _IntervalNode) -> _IntervalNode:
    _refresh_interval_node(node)
    balance = _interval_height(node.left) - _interval_height(node.right)
    if balance > 1:
        if node.left is not None and _interval_height(node.left.left) < _interval_height(
            node.left.right
        ):
            node.left = _rotate_interval_left(node.left)
        return _rotate_interval_right(node)
    if balance < -1:
        if node.right is not None and _interval_height(node.right.right) < _interval_height(
            node.right.left
        ):
            node.right = _rotate_interval_right(node.right)
        return _rotate_interval_left(node)
    return node


def _insert_interval_node(
    node: _IntervalNode | None,
    inserted: _IntervalNode,
) -> _IntervalNode:
    if node is None:
        return inserted
    if inserted.key < node.key:
        node.left = _insert_interval_node(node.left, inserted)
    else:
        node.right = _insert_interval_node(node.right, inserted)
    return _rebalance_interval_node(node)


def _delete_interval_node(
    node: _IntervalNode | None,
    key: tuple[int, int],
) -> _IntervalNode | None:
    if node is None:
        return None
    if key < node.key:
        node.left = _delete_interval_node(node.left, key)
    elif key > node.key:
        node.right = _delete_interval_node(node.right, key)
    elif node.left is None:
        return node.right
    elif node.right is None:
        return node.left
    else:
        successor = node.right
        while successor.left is not None:
            successor = successor.left
        node.key = successor.key
        node.end = successor.end
        node.item = successor.item
        node.right = _delete_interval_node(node.right, successor.key)
    return _rebalance_interval_node(node)


def _collect_interval_overlaps(
    node: _IntervalNode | None,
    lower: int,
    upper: int,
    result: list[int],
) -> None:
    if node is None or node.max_end < lower:
        return
    if node.left is not None and node.left.max_end >= lower:
        _collect_interval_overlaps(node.left, lower, upper, result)
    if node.key[0] <= upper and node.end >= lower:
        result.append(node.item)
    if node.key[0] <= upper:
        _collect_interval_overlaps(node.right, lower, upper, result)


class _AxisIntervalIndex:
    """Dynamic interval reporting plus constant-logarithmic overlap counts."""

    def __init__(self, intervals: Sequence[tuple[int, int, int]]) -> None:
        self._coordinates = tuple(
            sorted({coordinate for _item, start, end in intervals for coordinate in (start, end)})
        )
        self._starts = _Fenwick(len(self._coordinates))
        self._ends = _Fenwick(len(self._coordinates))
        self._intervals: dict[int, tuple[int, int]] = {}
        self._root: _IntervalNode | None = None
        for item, start, end in intervals:
            self.insert(item, start, end)

    def insert(self, item: int, start: int, end: int) -> None:
        if item in self._intervals:
            raise RuntimeError("rectangle interval is already indexed")
        self._intervals[item] = (start, end)
        self._starts.add(bisect_left(self._coordinates, start), 1)
        self._ends.add(bisect_left(self._coordinates, end), 1)
        self._root = _insert_interval_node(
            self._root,
            _IntervalNode((start, item), end, item),
        )

    def remove(self, item: int) -> None:
        start, end = self._intervals.pop(item)
        self._starts.add(bisect_left(self._coordinates, start), -1)
        self._ends.add(bisect_left(self._coordinates, end), -1)
        self._root = _delete_interval_node(self._root, (start, item))

    def overlap_count(self, lower: int, upper: int) -> int:
        starts_through_upper = self._starts.prefix(bisect_right(self._coordinates, upper))
        ends_before_lower = self._ends.prefix(bisect_left(self._coordinates, lower))
        return starts_through_upper - ends_before_lower

    def overlapping(self, lower: int, upper: int) -> tuple[int, ...]:
        result: list[int] = []
        _collect_interval_overlaps(self._root, lower, upper, result)
        return tuple(result)


class _RectangleGrid:
    """Dynamic sparse 2D buckets for small rectangle queries."""

    def __init__(self, rectangles: Mapping[int, Rect]) -> None:
        self._buckets: dict[tuple[int, int], set[int]] = defaultdict(set)
        self._buckets_by_item: dict[int, tuple[tuple[int, int], ...]] = {}
        self._wide: set[int] = set()
        for item, rect in rectangles.items():
            self.insert(item, rect)

    def insert(self, item: int, rect: Rect) -> None:
        buckets = _rectangle_grid_buckets(rect)
        if buckets is None:
            self._wide.add(item)
            self._buckets_by_item[item] = ()
            return
        self._buckets_by_item[item] = buckets
        for bucket in buckets:
            self._buckets[bucket].add(item)

    def remove(self, item: int) -> None:
        buckets = self._buckets_by_item.pop(item)
        if not buckets:
            self._wide.discard(item)
            return
        for bucket in buckets:
            self._buckets[bucket].remove(item)
            if not self._buckets[bucket]:
                del self._buckets[bucket]

    def candidates(self, rect: Rect) -> set[int] | None:
        buckets = _rectangle_grid_buckets(rect)
        if buckets is None:
            return None
        result = set(self._wide)
        for bucket in buckets:
            result.update(self._buckets.get(bucket, ()))
        return result


def _rectangle_grid_buckets(rect: Rect) -> tuple[tuple[int, int], ...] | None:
    first_row = (rect.row_min - 1) // _RECT_GRID_BUCKET_SIZE
    last_row = (rect.row_max - 1) // _RECT_GRID_BUCKET_SIZE
    first_col = (rect.col_min - 1) // _RECT_GRID_BUCKET_SIZE
    last_col = (rect.col_max - 1) // _RECT_GRID_BUCKET_SIZE
    row_count = last_row - first_row + 1
    col_count = last_col - first_col + 1
    if row_count * col_count > _RECT_GRID_MAX_CELLS:
        return None
    return tuple(
        (row_bucket, col_bucket)
        for row_bucket in range(first_row, last_row + 1)
        for col_bucket in range(first_col, last_col + 1)
    )


class _SpatialRectangleNode:
    """Spatially partitioned bounds with a rect-key lower bound."""

    __slots__ = ("bounds", "item", "left", "min_order", "right")

    def __init__(
        self,
        bounds: Rect,
        min_order: int,
        *,
        item: int | None = None,
        left: _SpatialRectangleNode | None = None,
        right: _SpatialRectangleNode | None = None,
    ) -> None:
        self.bounds = bounds
        self.min_order = min_order
        self.item = item
        self.left = left
        self.right = right


def _build_spatial_rectangle_tree(
    rectangles: Sequence[Rect],
    items: Sequence[int],
    order_by_item: Mapping[int, int],
) -> _SpatialRectangleNode | None:
    if not items:
        return None
    bounds = _rectangles_bounds(rectangles[item] for item in items)
    if len(items) == 1:
        item = items[0]
        return _SpatialRectangleNode(
            bounds,
            order_by_item[item],
            item=item,
        )
    row_span = bounds.row_max - bounds.row_min
    col_span = bounds.col_max - bounds.col_min
    if col_span > row_span:
        ordered = sorted(
            items,
            key=lambda item: (
                rectangles[item].col_min + rectangles[item].col_max,
                rectangles[item].row_min + rectangles[item].row_max,
                *_rect_key(rectangles[item]),
                item,
            ),
        )
    else:
        ordered = sorted(
            items,
            key=lambda item: (
                rectangles[item].row_min + rectangles[item].row_max,
                rectangles[item].col_min + rectangles[item].col_max,
                *_rect_key(rectangles[item]),
                item,
            ),
        )
    middle = len(ordered) // 2
    left = _build_spatial_rectangle_tree(rectangles, ordered[:middle], order_by_item)
    right = _build_spatial_rectangle_tree(rectangles, ordered[middle:], order_by_item)
    if left is None:
        return right
    if right is None:
        return left
    return _SpatialRectangleNode(
        _bounding_rect(left.bounds, right.bounds),
        min(left.min_order, right.min_order),
        left=left,
        right=right,
    )


def _iter_spatial_ordered_intersections(
    root: _SpatialRectangleNode | None,
    query: Rect,
) -> Iterable[int]:
    if root is None or not _rectangles_intersect(root.bounds, query):
        return
    pending: list[tuple[int, _SpatialRectangleNode]] = [(root.min_order, root)]
    while pending:
        _order, node = heappop(pending)
        if node.item is not None:
            yield node.item
            continue
        if node.left is not None and _rectangles_intersect(node.left.bounds, query):
            heappush(pending, (node.left.min_order, node.left))
        if node.right is not None and _rectangles_intersect(node.right.bounds, query):
            heappush(pending, (node.right.min_order, node.right))


class _RectangleIndex:
    """Sparse dynamic rectangle index used by barriers and closure expansion."""

    def __init__(self, rectangles: Sequence[Rect]) -> None:
        self._rectangles = {index: rect for index, rect in enumerate(rectangles)}
        self._initial_rectangles = tuple(rectangles)
        self._spatial_root: _SpatialRectangleNode | None = None
        self._spatial_built = False
        self._rows = _AxisIntervalIndex(
            tuple((index, rect.row_min, rect.row_max) for index, rect in self._rectangles.items())
        )
        self._columns = _AxisIntervalIndex(
            tuple((index, rect.col_min, rect.col_max) for index, rect in self._rectangles.items())
        )
        self._grid = _RectangleGrid(self._rectangles)

    def insert(self, item: int, rect: Rect) -> None:
        if item in self._rectangles:
            raise RuntimeError("rectangle is already indexed")
        self._rectangles[item] = rect
        self._rows.insert(item, rect.row_min, rect.row_max)
        self._columns.insert(item, rect.col_min, rect.col_max)
        self._grid.insert(item, rect)

    def remove(self, item: int) -> None:
        self._rows.remove(item)
        self._columns.remove(item)
        self._grid.remove(item)
        del self._rectangles[item]

    def intersections(self, rect: Rect, *, exclude: int | None = None) -> tuple[int, ...]:
        row_count = self._rows.overlap_count(rect.row_min, rect.row_max)
        column_count = self._columns.overlap_count(rect.col_min, rect.col_max)
        grid_candidates = self._grid.candidates(rect)
        minimum_axis_count = min(row_count, column_count)
        maximum_axis_count = max(row_count, column_count)
        if grid_candidates is not None and len(grid_candidates) < minimum_axis_count:
            candidates: Iterable[int] = grid_candidates
        elif (
            minimum_axis_count >= _RECT_BALANCED_AXIS_MIN
            and maximum_axis_count <= minimum_axis_count * _RECT_BALANCED_AXIS_RATIO
        ):
            row_candidates = set(self._rows.overlapping(rect.row_min, rect.row_max))
            column_candidates = set(self._columns.overlapping(rect.col_min, rect.col_max))
            candidates = row_candidates & column_candidates
        elif row_count <= column_count:
            candidates = self._rows.overlapping(rect.row_min, rect.row_max)
        else:
            candidates = self._columns.overlapping(rect.col_min, rect.col_max)
        return tuple(
            sorted(
                (
                    item
                    for item in candidates
                    if item != exclude and _rectangles_intersect(rect, self._rectangles[item])
                ),
                key=lambda item: (*_rect_key(self._rectangles[item]), item),
            )
        )

    def intersects_any(self, rect: Rect) -> bool:
        return bool(self.intersections(rect))

    def ordered_intersections(self, rect: Rect) -> Iterable[int]:
        """Yield initial rectangles in rect-key order through a spatial tree."""
        if not self._spatial_built:
            ordered_items = tuple(
                sorted(
                    range(len(self._initial_rectangles)),
                    key=lambda item: (
                        *_rect_key(self._initial_rectangles[item]),
                        item,
                    ),
                )
            )
            order_by_item = {item: order for order, item in enumerate(ordered_items)}
            self._spatial_root = _build_spatial_rectangle_tree(
                self._initial_rectangles,
                ordered_items,
                order_by_item,
            )
            self._spatial_built = True
        return _iter_spatial_ordered_intersections(self._spatial_root, rect)

    def rectangle(self, item: int) -> Rect:
        return self._rectangles[item]

    def rectangles(self) -> tuple[Rect, ...]:
        return tuple(self._rectangles.values())


def _rectangles_intersect(left: Rect, right: Rect) -> bool:
    return not (
        left.row_max < right.row_min
        or right.row_max < left.row_min
        or left.col_max < right.col_min
        or right.col_max < left.col_min
    )


class _RegionLocator:
    """Locate cells in pairwise-disjoint rectangles during an ordered scan."""

    def __init__(self, seeds: Sequence[_Seed]) -> None:
        self._ordered = sorted(enumerate(seeds), key=lambda item: (item[1].rect.row_min, item[0]))
        self._next = 0
        self._active: list[tuple[int, _Seed]] = []
        self._active_starts: list[int] = []
        self._row = 0

    def locate(self, row: int, col: int) -> int | None:
        if row < self._row:
            raise ValueError("cell stream must be strictly ordered without duplicates")
        if row != self._row:
            self._row = row
            self._active = [item for item in self._active if item[1].rect.row_max >= row]
            while self._next < len(self._ordered):
                item = self._ordered[self._next]
                if item[1].rect.row_min > row:
                    break
                self._next += 1
                if item[1].rect.row_max >= row:
                    self._active.append(item)
            self._active.sort(key=lambda item: (item[1].rect.col_min, item[1].rect.col_max))
            self._active_starts = [item[1].rect.col_min for item in self._active]
        candidate_index = bisect_right(self._active_starts, col) - 1
        if candidate_index < 0:
            return None
        seed_index, seed = self._active[candidate_index]
        return seed_index if contains(seed.rect, row, col) else None


def analyze_sheet_regions(
    summary: SheetParseSummary,
    styles: StyleCatalog,
    cell_stream_factory: CellStreamFactory,
    options: RegionOptions | None = None,
    *,
    coordinate_stream_factory: CellStreamFactory | None = None,
    header_cell_provider: HeaderCellProvider | None = None,
    column_profile_provider: ColumnProfileProvider | None = None,
) -> RegionAnalysis:
    """Detect and profile all regions in one parsed worksheet.

    ``cell_stream_factory`` must return a fresh coordinate-ordered iterable on
    every call.  Analysis performs bounded sparse passes rather than retaining
    worksheet values in a dense structure.
    """
    active_options = options or RegionOptions()
    tables = _validated_tables(summary.tables)
    _validate_merges(summary.merges, tables)
    dense_rect = _dense_origin_rect(summary, tables)
    heuristic_rects = (
        (dense_rect,)
        if dense_rect is not None
        else _detect_heuristic_rects(
            (coordinate_stream_factory or cell_stream_factory)(),
            tables=tables,
            merges=summary.merges,
            gap_tol=active_options.gap_tol,
        )
    )
    seeds = [*(_Seed(_table_rect(table), "table", table) for table in tables)]
    seeds.extend(_Seed(rect, "region") for rect in heuristic_rects)
    seeds.sort(key=_seed_sort_key)
    _require_disjoint(seeds)

    header_cells = (
        cell_stream_factory()
        if header_cell_provider is None
        else header_cell_provider(tuple(seed.rect for seed in seeds if seed.kind != "table"))
    )
    decisions = _infer_headers(
        seeds,
        summary.merges,
        styles,
        header_cells,
        threshold=active_options.header_threshold,
    )
    if column_profile_provider is None:
        profiles = _profile_columns(
            seeds,
            decisions,
            cell_stream_factory(),
            cell_count=summary.cell_count,
            options=active_options,
        )
    else:
        requests = tuple(
            ColumnProfileRequest(
                seed.rect,
                decisions[index].rows,
                decisions[index].headers,
                decisions[index].normalized_headers,
                0 if seed.table is None else seed.table.totals_rows,
            )
            for index, seed in enumerate(seeds)
        )
        profiles = column_profile_provider(requests, summary.cell_count, active_options)
        if len(profiles) != len(requests) or any(
            len(profile) != _width(request.rect)
            for request, profile in zip(requests, profiles, strict=True)
        ):
            raise ValueError("column profile provider returned an invalid result shape")
    regions = tuple(
        DetectedRegion(
            n=index,
            rect=seed.rect,
            header_rows=decisions[index].rows,
            kind=seed.kind,
            list_object_name=None if seed.table is None else seed.table.name,
            confidence=decisions[index].confidence,
            columns=profiles[index],
        )
        for index, seed in enumerate(seeds)
    )
    warnings: tuple[RegionWarning, ...] = ()
    if summary.cell_count > active_options.large_sheet_threshold:
        stride = max(2, math.ceil(summary.cell_count / active_options.large_sheet_threshold))
        warnings = (
            RegionWarning(
                code="W_LARGE_SHEET",
                ref=sheet_symbol_id(summary.descriptor.name),
                message="Large worksheet was indexed with reduced-rate dtype sampling.",
                related={
                    "cellCount": summary.cell_count,
                    "dtypeSampleStride": stride,
                    "dtypeSampleLimit": active_options.large_dtype_sample_limit,
                },
            ),
        )
    return RegionAnalysis(regions=regions, warnings=warnings)


def _dense_origin_rect(
    summary: SheetParseSummary,
    tables: Sequence[TableInfo],
) -> Rect | None:
    """Return the exact A1-origin rectangle certified by the parsed cell count."""
    if tables or summary.merges or summary.max_row < 1 or summary.max_col < 1:
        return None
    if summary.cell_count != summary.max_row * summary.max_col:
        return None
    return Rect(1, summary.max_row, 1, summary.max_col)


def _validated_tables(tables: Sequence[TableInfo]) -> tuple[TableInfo, ...]:
    parsed = tuple((table, _table_rect(table)) for table in tables)
    ordered_pairs = tuple(sorted(parsed, key=lambda item: (*_rect_key(item[1]), item[0].name)))
    ordered = tuple(table for table, _rect in ordered_pairs)
    rects = tuple(rect for _table, rect in ordered_pairs)
    overlaps_with_prior = {
        current_index for _prior_index, current_index in _overlapping_pairs(rects)
    }
    names: set[str] = set()
    for index, (table, rect) in enumerate(ordered_pairs):
        width = rect.col_max - rect.col_min + 1
        height = rect.row_max - rect.row_min + 1
        if not table.name or not table.display_name or any(not name for name in table.columns):
            _table_corrupt(table, "has an empty name or column name")
        if table.name.casefold() in names:
            _table_corrupt(table, "duplicates another ListObject name")
        names.add(table.name.casefold())
        if len(table.columns) != width:
            _table_corrupt(table, "column count does not match its declared range")
        if table.header_rows + table.totals_rows > height:
            _table_corrupt(table, "header and totals rows exceed its declared range")
        if index in overlaps_with_prior:
            _table_corrupt(table, "overlaps another ListObject")
    return ordered


def _table_rect(table: TableInfo) -> Rect:
    try:
        return parse_rect(table.ref)
    except ValueError:
        _table_corrupt(table, "has an invalid declared range")


def _table_corrupt(table: TableInfo, reason: str) -> NoReturn:
    raise ExcelLSPError(
        ErrorCode.CORRUPT,
        f"Excel Table {table.name!r} {reason}.",
        details={"table": table.name, "ref": table.ref},
    )


def _validate_merges(merges: Sequence[Rect], tables: Sequence[TableInfo]) -> None:
    ordered = tuple(sorted(merges, key=_rect_key))
    tagged: list[tuple[Rect, bool, int]] = [
        (merge, False, index) for index, merge in enumerate(ordered)
    ]
    tagged.extend((_table_rect(table), True, index) for index, table in enumerate(tables))
    tagged.sort(key=lambda item: (*_rect_key(item[0]), item[1], item[2]))
    merge_table_overlaps: set[int] = set()
    merge_prior_overlaps: set[int] = set()
    for left_position, right_position in _overlapping_pairs(tuple(item[0] for item in tagged)):
        _left_rect, left_is_table, left_index = tagged[left_position]
        _right_rect, right_is_table, right_index = tagged[right_position]
        if left_is_table != right_is_table:
            merge_table_overlaps.add(right_index if left_is_table else left_index)
        elif not left_is_table:
            merge_prior_overlaps.add(max(left_index, right_index))

    for index in range(len(ordered)):
        if index in merge_table_overlaps:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "A merged range overlaps an Excel Table.",
            )
        if index in merge_prior_overlaps:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Worksheet contains overlapping merged ranges.",
            )


def _detect_heuristic_rects(
    cells: Iterable[RegionCell],
    *,
    tables: Sequence[TableInfo],
    merges: Sequence[Rect],
    gap_tol: int,
) -> tuple[Rect, ...]:
    table_rects = tuple(_table_rect(table) for table in tables)
    table_barriers = _RectangleIndex(table_rects)
    table_locator = _RegionLocator(
        tuple(_Seed(rect, "table", table) for table, rect in zip(tables, table_rects, strict=True))
    )
    merge_by_anchor = {(merge.row_min, merge.col_min): merge for merge in merges}
    anchored_merges: list[Rect] = []
    exact_runs: list[_Run] = []
    previous: tuple[int, int] | None = None
    current: _Run | None = None

    for cell in cells:
        coordinate = (cell.row, cell.col)
        if previous is not None and coordinate <= previous:
            raise ValueError("cell stream must be strictly ordered without duplicates")
        previous = coordinate
        merge = merge_by_anchor.get(coordinate)
        if merge is not None:
            anchored_merges.append(merge)
        if table_locator.locate(cell.row, cell.col) is not None:
            continue
        if current is not None and current.row == cell.row and current.col_max + 1 == cell.col:
            current = _Run(current.row, current.col_min, cell.col)
            continue
        if current is not None:
            exact_runs.append(current)
        current = _Run(cell.row, cell.col, cell.col)
    if current is not None:
        exact_runs.append(current)

    runs = _coalesce_row_runs(
        exact_runs,
        gap_tol=gap_tol,
        table_barriers=table_barriers,
    )
    rectangles: list[Rect] = []
    if anchored_merges:
        spans = tuple(
            _Span(
                merge.row_min,
                merge.row_max,
                merge.col_min,
                merge.col_max,
            )
            for merge in anchored_merges
        )
        rectangles.extend(
            _component_first_sparse_rectangles(
                runs,
                spans,
                table_rects,
                table_barriers,
                gap_tol=gap_tol,
            )
        )
    else:
        zones = _partition_runs_around_tables(runs, table_rects, table_barriers)
        for zone in zones:
            rectangles.extend(_component_rectangles(zone, gap_tol=gap_tol))
    return tuple(sorted(rectangles, key=_rect_key))


def _component_first_sparse_rectangles(
    runs: Sequence[_Run],
    spans: Sequence[_Span],
    tables: Sequence[Rect],
    table_barriers: _RectangleIndex,
    *,
    gap_tol: int,
) -> tuple[Rect, ...]:
    pending = deque(
        _closed_sparse_components(
            runs,
            spans,
            table_barriers,
            gap_tol=gap_tol,
        )
    )
    rectangles: list[Rect] = []
    while pending:
        component = pending.popleft()
        ordered_candidates = iter(table_barriers.ordered_intersections(component.bounds))
        table_index = next(ordered_candidates, None)
        if table_index is None:
            rectangles.append(component.bounds)
            continue

        zone = _SparseZone.create(
            None if not component.runs else _RunView.complete(component.runs),
            component.spans,
        )
        if zone is None:
            continue
        table = tables[table_index]
        if table.row_min <= component.bounds.row_min and table.row_max >= component.bounds.row_max:
            full_height_tables = [table]
            for candidate in ordered_candidates:
                candidate_table = table_barriers.rectangle(candidate)
                if (
                    candidate_table.row_min > component.bounds.row_min
                    or candidate_table.row_max < component.bounds.row_max
                ):
                    break
                full_height_tables.append(candidate_table)
            children = _partition_full_height_tables(
                zone,
                tuple(
                    sorted(
                        full_height_tables,
                        key=lambda rect: (
                            rect.col_min,
                            rect.col_max,
                            *_rect_key(rect),
                        ),
                    )
                ),
            )
        else:
            children = _split_sparse_zone_around_table(zone, table)

        for child in children:
            child_runs = () if child.runs is None else child.runs.materialize()
            pending.extend(
                _closed_sparse_components(
                    child_runs,
                    child.spans,
                    table_barriers,
                    gap_tol=gap_tol,
                )
            )
    return tuple(sorted(rectangles, key=_rect_key))


def _coalesce_row_runs(
    runs: Sequence[_Run],
    *,
    gap_tol: int,
    table_barriers: _RectangleIndex,
) -> tuple[_Run, ...]:
    result: list[_Run] = []
    for run in sorted(runs, key=lambda item: (item.row, item.col_min, item.col_max)):
        if not result or result[-1].row != run.row:
            result.append(run)
            continue
        prior = result[-1]
        blank_columns = run.col_min - prior.col_max - 1
        should_merge = run.col_min <= prior.col_max + 1
        if not should_merge and blank_columns <= gap_tol:
            bridge = Rect(run.row, run.row, prior.col_max + 1, run.col_min - 1)
            bridge_tables = tuple(
                table_barriers.rectangle(table_index)
                for table_index in table_barriers.intersections(bridge)
            )
            should_merge = not any(
                table.col_min <= run.col_min - 1 and table.col_max >= prior.col_max + 1
                for table in bridge_tables
            )
        if should_merge:
            result[-1] = _Run(run.row, prior.col_min, max(prior.col_max, run.col_max))
        else:
            result.append(run)
    return tuple(result)


def _partition_runs_around_tables(
    runs: Sequence[_Run],
    tables: Sequence[Rect],
    table_barriers: _RectangleIndex,
) -> tuple[tuple[_Run, ...], ...]:
    zones = _partition_sparse_geometry_around_tables(
        runs,
        (),
        tables,
        table_barriers,
    )
    return tuple(zone.runs.materialize() for zone in zones if zone.runs is not None)


def _partition_sparse_geometry_around_tables(
    runs: Sequence[_Run],
    spans: Sequence[_Span],
    tables: Sequence[Rect],
    table_barriers: _RectangleIndex,
) -> tuple[_SparseZone, ...]:
    run_view = None if not runs else _RunView.complete(runs)
    initial = _SparseZone.create(run_view, spans)
    if initial is None:
        return ()
    pending = deque((initial,))
    zones: list[_SparseZone] = []
    while pending:
        zone = pending.popleft()
        ordered_candidates = iter(table_barriers.ordered_intersections(zone.bounds))
        table_index = next(ordered_candidates, None)
        if table_index is None:
            zones.append(zone)
            continue
        table = tables[table_index]
        if table.row_min <= zone.bounds.row_min and table.row_max >= zone.bounds.row_max:
            full_height_tables = [table]
            for candidate in ordered_candidates:
                candidate_table = table_barriers.rectangle(candidate)
                if (
                    candidate_table.row_min > zone.bounds.row_min
                    or candidate_table.row_max < zone.bounds.row_max
                ):
                    break
                full_height_tables.append(candidate_table)
            pending.extend(
                _partition_full_height_tables(
                    zone,
                    tuple(
                        sorted(
                            full_height_tables,
                            key=lambda rect: (rect.col_min, rect.col_max, *_rect_key(rect)),
                        )
                    ),
                )
            )
            continue
        pending.extend(_split_sparse_zone_around_table(zone, table))
    return tuple(zones)


def _split_sparse_zone_around_table(
    zone: _SparseZone,
    table: Rect,
) -> tuple[_SparseZone, ...]:
    run_children: list[_RunView | None] = [None, None, None, None]
    if zone.runs is not None:
        middle_start = zone.runs.source.lower_bound(
            zone.runs.start,
            zone.runs.stop,
            table.row_min,
        )
        middle_stop = zone.runs.source.upper_bound(
            middle_start,
            zone.runs.stop,
            table.row_max,
        )
        if zone.runs.start < middle_start:
            run_children[0] = _RunView(
                zone.runs.source,
                zone.runs.start,
                middle_start,
            )
        if middle_stop < zone.runs.stop:
            run_children[1] = _RunView(
                zone.runs.source,
                middle_stop,
                zone.runs.stop,
            )
        left: list[_Run] = []
        right: list[_Run] = []
        for run in zone.runs.source.runs[middle_start:middle_stop]:
            if run.col_min < table.col_min:
                left.append(
                    _Run(
                        run.row,
                        run.col_min,
                        min(run.col_max, table.col_min - 1),
                    )
                )
            if run.col_max > table.col_max:
                right.append(
                    _Run(
                        run.row,
                        max(run.col_min, table.col_max + 1),
                        run.col_max,
                    )
                )
        if left:
            run_children[2] = _RunView.complete(left)
        if right:
            run_children[3] = _RunView.complete(right)

    span_children: tuple[list[_Span], ...] = ([], [], [], [])
    for span in zone.spans:
        _split_span_around_table(span, table, span_children)
    children: list[_SparseZone] = []
    for child_runs, child_spans in zip(
        run_children,
        span_children,
        strict=True,
    ):
        child = _SparseZone.create(child_runs, child_spans)
        if child is not None:
            children.append(child)
    return tuple(children)


def _partition_full_height_tables(
    zone: _SparseZone,
    tables: Sequence[Rect],
) -> tuple[_SparseZone, ...]:
    bands: list[tuple[int, int]] = []
    cursor = zone.bounds.col_min
    for table in tables:
        table_min = max(zone.bounds.col_min, table.col_min)
        table_max = min(zone.bounds.col_max, table.col_max)
        if cursor < table_min:
            bands.append((cursor, table_min - 1))
        cursor = max(cursor, table_max + 1)
    if cursor <= zone.bounds.col_max:
        bands.append((cursor, zone.bounds.col_max))
    if not bands:
        return ()

    band_maxima = tuple(col_max for _col_min, col_max in bands)
    grouped_runs: list[list[_Run]] = [[] for _band in bands]
    if zone.runs is not None:
        for run in zone.runs.iter_runs():
            band_index = bisect_left(band_maxima, run.col_min)
            while band_index < len(bands) and bands[band_index][0] <= run.col_max:
                band_min, band_max = bands[band_index]
                fragment_min = max(run.col_min, band_min)
                fragment_max = min(run.col_max, band_max)
                if fragment_min <= fragment_max:
                    grouped_runs[band_index].append(_Run(run.row, fragment_min, fragment_max))
                band_index += 1

    grouped_spans: list[list[_Span]] = [[] for _band in bands]
    for span in zone.spans:
        band_index = bisect_left(band_maxima, span.col_min)
        while band_index < len(bands) and bands[band_index][0] <= span.col_max:
            band_min, band_max = bands[band_index]
            fragment_min = max(span.col_min, band_min)
            fragment_max = min(span.col_max, band_max)
            if fragment_min <= fragment_max:
                grouped_spans[band_index].append(
                    _Span(
                        span.row_min,
                        span.row_max,
                        fragment_min,
                        fragment_max,
                    )
                )
            band_index += 1

    children: list[_SparseZone] = []
    for run_group, span_group in zip(grouped_runs, grouped_spans, strict=True):
        child_runs = None if not run_group else _RunView.complete(run_group)
        child = _SparseZone.create(child_runs, span_group)
        if child is not None:
            children.append(child)
    return tuple(children)


def _split_span_around_table(
    span: _Span,
    table: Rect,
    groups: tuple[list[_Span], ...],
) -> None:
    above, below, left, right = groups
    if span.row_min < table.row_min:
        above.append(
            _Span(
                span.row_min,
                min(span.row_max, table.row_min - 1),
                span.col_min,
                span.col_max,
            )
        )
    if span.row_max > table.row_max:
        below.append(
            _Span(
                max(span.row_min, table.row_max + 1),
                span.row_max,
                span.col_min,
                span.col_max,
            )
        )
    middle_min = max(span.row_min, table.row_min)
    middle_max = min(span.row_max, table.row_max)
    if middle_min > middle_max:
        return
    if span.col_min < table.col_min:
        left.append(
            _Span(
                middle_min,
                middle_max,
                span.col_min,
                min(span.col_max, table.col_min - 1),
            )
        )
    if span.col_max > table.col_max:
        right.append(
            _Span(
                middle_min,
                middle_max,
                max(span.col_min, table.col_max + 1),
                span.col_max,
            )
        )


def _component_rectangles(runs: Sequence[_Run], *, gap_tol: int) -> tuple[Rect, ...]:
    if not runs:
        return ()
    ordered = tuple(sorted(runs, key=lambda run: (run.row, run.col_min, run.col_max)))
    disjoint = _DisjointSet(len(ordered))
    indices_by_row: dict[int, list[int]] = defaultdict(list)
    for index, run in enumerate(ordered):
        indices_by_row[run.row].append(index)
    active_rows: deque[int] = deque()
    for row in sorted(indices_by_row):
        while active_rows and row - active_rows[0] - 1 > gap_tol:
            active_rows.popleft()
        for prior_row in active_rows:
            _union_near_runs(
                ordered,
                indices_by_row[row],
                indices_by_row[prior_row],
                disjoint,
                gap_tol=gap_tol,
            )
        active_rows.append(row)
    grouped: dict[int, list[_Run]] = defaultdict(list)
    for index, run in enumerate(ordered):
        grouped[disjoint.find(index)].append(run)
    rectangles = [_runs_rect(component) for component in grouped.values()]
    return _merge_intersecting_rectangles(rectangles)


def _closed_sparse_components(
    runs: Sequence[_Run],
    spans: Sequence[_Span],
    table_barriers: _RectangleIndex,
    *,
    gap_tol: int,
) -> tuple[_SparseComponent, ...]:
    """Return table-aware proximity components after bounding-box closure.

    The sweep keeps every old-old witness edge internal and every root bounding
    box disjoint; root-local binomial spatial blocks retain members without
    copying a growing component on each union.
    """
    ordered_runs = tuple(
        sorted(
            runs,
            key=lambda run: (run.row, run.col_min, run.col_max),
        )
    )
    ordered_spans = tuple(
        sorted(
            spans,
            key=lambda span: (
                span.row_min,
                span.col_min,
                span.row_max,
                span.col_max,
            ),
        )
    )
    run_count = len(ordered_runs)
    primitives = tuple(
        (
            *(
                _SparsePrimitive(
                    index,
                    Rect(run.row, run.row, run.col_min, run.col_max),
                )
                for index, run in enumerate(ordered_runs)
            ),
            *(
                _SparsePrimitive(
                    run_count + index,
                    span.rect,
                )
                for index, span in enumerate(ordered_spans)
            ),
        )
    )
    if not primitives:
        return ()

    rectangles = tuple(primitive.rect for primitive in primitives)
    components = _DisjointSet(len(primitives))
    component_blocks: list[tuple[_SparseMemberBlock, ...]] = [() for _primitive in primitives]
    component_bounds = list(rectangles)
    component_index = _RectangleIndex(rectangles)
    active_component_index = _RectangleIndex(rectangles)
    for item in range(len(primitives)):
        component_index.remove(item)
        active_component_index.remove(item)
    active_roots: set[int] = set()
    expiration_heap: list[tuple[int, int]] = []
    corridor_cache: dict[tuple[str, int, int, int, int], bool] = {}
    ordered_primitive_ids = tuple(
        sorted(
            range(len(primitives)),
            key=lambda item: (
                *_rect_key(rectangles[item]),
                item,
            ),
        )
    )

    # The witness graph is static.  Before each later primitive, every old-old
    # witness edge is internal to a root and active root bounds are disjoint.
    # Evaluate only the new primitive's old-neighbor edges, then bounding-box
    # close the enlarged root to restore that invariant.
    for item in ordered_primitive_ids:
        _expire_sparse_components(
            rectangles[item].row_min,
            gap_tol,
            components,
            component_bounds,
            active_component_index,
            active_roots,
            expiration_heap,
        )
        component_blocks[item] = (_build_sparse_member_block((item,), rectangles),)
        component_index.insert(item, rectangles[item])
        active_component_index.insert(item, rectangles[item])
        active_roots.add(item)
        heappush(expiration_heap, (rectangles[item].row_max, item))
        root = item
        initial_targets = _intersecting_sparse_components(
            active_component_index,
            rectangles[item],
            exclude=item,
        )
        if initial_targets:
            root = _merge_sparse_component_roots(
                root,
                initial_targets,
                components,
                component_blocks,
                component_bounds,
                rectangles,
                component_index,
                active_component_index,
                active_roots,
                expiration_heap,
            )
            root = _close_sparse_component_bounds(
                root,
                components,
                component_blocks,
                component_bounds,
                rectangles,
                component_index,
                active_component_index,
                active_roots,
                expiration_heap,
            )

        near_roots = _nearby_sparse_components(
            active_component_index,
            rectangles[item],
            gap_tol=gap_tol,
            exclude=root,
        )
        if not near_roots:
            continue

        successful_roots = tuple(
            target
            for target in near_roots
            if _sparse_root_has_witness(
                item,
                component_blocks[target],
                rectangles,
                rectangles[item],
                gap_tol=gap_tol,
                table_barriers=table_barriers,
                corridor_cache=corridor_cache,
            )
        )
        if successful_roots:
            root = _merge_sparse_component_roots(
                root,
                successful_roots,
                components,
                component_blocks,
                component_bounds,
                rectangles,
                component_index,
                active_component_index,
                active_roots,
                expiration_heap,
            )
            _close_sparse_component_bounds(
                root,
                components,
                component_blocks,
                component_bounds,
                rectangles,
                component_index,
                active_component_index,
                active_roots,
                expiration_heap,
            )

    result: list[_SparseComponent] = []
    for item in range(len(primitives)):
        if components.find(item) != item:
            continue
        member_ids = tuple(
            sorted(member_id for block in component_blocks[item] for member_id in block.member_ids)
        )
        component_runs = tuple(
            ordered_runs[member_id] for member_id in member_ids if member_id < run_count
        )
        component_spans = tuple(
            ordered_spans[member_id - run_count]
            for member_id in member_ids
            if member_id >= run_count
        )
        result.append(
            _SparseComponent(
                component_runs,
                component_spans,
                component_bounds[item],
            )
        )
    return tuple(
        sorted(
            result,
            key=lambda component: (
                *_rect_key(component.bounds),
                tuple((run.row, run.col_min, run.col_max) for run in component.runs),
                tuple(
                    (
                        span.row_min,
                        span.col_min,
                        span.row_max,
                        span.col_max,
                    )
                    for span in component.spans
                ),
            ),
        )
    )


def _close_sparse_component_bounds(
    root: int,
    components: _DisjointSet,
    component_blocks: list[tuple[_SparseMemberBlock, ...]],
    component_bounds: list[Rect],
    rectangles: Sequence[Rect],
    component_index: _RectangleIndex,
    active_component_index: _RectangleIndex,
    active_roots: set[int],
    expiration_heap: list[tuple[int, int]],
) -> int:
    while True:
        targets = _intersecting_sparse_components(
            component_index,
            component_bounds[root],
            exclude=root,
        )
        if not targets:
            return root
        root = _merge_sparse_component_roots(
            root,
            targets,
            components,
            component_blocks,
            component_bounds,
            rectangles,
            component_index,
            active_component_index,
            active_roots,
            expiration_heap,
        )


def _merge_sparse_component_roots(
    root: int,
    targets: Sequence[int],
    components: _DisjointSet,
    component_blocks: list[tuple[_SparseMemberBlock, ...]],
    component_bounds: list[Rect],
    rectangles: Sequence[Rect],
    component_index: _RectangleIndex,
    active_component_index: _RectangleIndex,
    active_roots: set[int],
    expiration_heap: list[tuple[int, int]],
) -> int:
    component_index.remove(root)
    if root in active_roots:
        active_component_index.remove(root)
        active_roots.remove(root)
    combined_bounds = component_bounds[root]
    block_groups = [component_blocks[root]]
    component_blocks[root] = ()
    for target in targets:
        component_index.remove(target)
        if target in active_roots:
            active_component_index.remove(target)
            active_roots.remove(target)
        combined_bounds = _bounding_rect(
            combined_bounds,
            component_bounds[target],
        )
        block_groups.append(component_blocks[target])
        component_blocks[target] = ()
        root = components.union(root, target)
    component_bounds[root] = combined_bounds
    component_blocks[root] = _meld_sparse_member_blocks(
        block_groups,
        rectangles,
    )
    component_index.insert(root, combined_bounds)
    active_component_index.insert(root, combined_bounds)
    active_roots.add(root)
    heappush(expiration_heap, (combined_bounds.row_max, root))
    return root


def _expire_sparse_components(
    next_row: int,
    gap_tol: int,
    components: _DisjointSet,
    component_bounds: Sequence[Rect],
    active_component_index: _RectangleIndex,
    active_roots: set[int],
    expiration_heap: list[tuple[int, int]],
) -> None:
    distance = gap_tol + 1
    while expiration_heap and expiration_heap[0][0] + distance < next_row:
        row_max, item = heappop(expiration_heap)
        root = components.find(item)
        if root != item or root not in active_roots or component_bounds[root].row_max != row_max:
            continue
        active_component_index.remove(root)
        active_roots.remove(root)


def _intersecting_sparse_components(
    component_index: _RectangleIndex,
    bounds: Rect,
    *,
    exclude: int,
) -> tuple[int, ...]:
    return component_index.intersections(bounds, exclude=exclude)


def _nearby_sparse_components(
    component_index: _RectangleIndex,
    bounds: Rect,
    *,
    gap_tol: int,
    exclude: int,
) -> tuple[int, ...]:
    return component_index.intersections(
        _expanded_proximity_rect(bounds, gap_tol=gap_tol),
        exclude=exclude,
    )


def _build_sparse_member_block(
    member_ids: tuple[int, ...],
    rectangles: Sequence[Rect],
) -> _SparseMemberBlock:
    return _SparseMemberBlock(
        member_ids,
        _RectangleIndex(tuple(rectangles[member_id] for member_id in member_ids)),
    )


def _meld_sparse_member_blocks(
    block_groups: Sequence[Sequence[_SparseMemberBlock]],
    rectangles: Sequence[Rect],
) -> tuple[_SparseMemberBlock, ...]:
    blocks_by_size: dict[int, _SparseMemberBlock] = {}
    for blocks in block_groups:
        for block in blocks:
            size = len(block.member_ids)
            while size in blocks_by_size:
                prior = blocks_by_size.pop(size)
                block = _build_sparse_member_block(
                    _merge_sorted_member_ids(
                        prior.member_ids,
                        block.member_ids,
                    ),
                    rectangles,
                )
                size *= 2
            blocks_by_size[size] = block
    return tuple(blocks_by_size[size] for size in sorted(blocks_by_size))


def _merge_sorted_member_ids(
    left: Sequence[int],
    right: Sequence[int],
) -> tuple[int, ...]:
    result: list[int] = []
    left_index = 0
    right_index = 0
    while left_index < len(left) and right_index < len(right):
        if left[left_index] <= right[right_index]:
            result.append(left[left_index])
            left_index += 1
        else:
            result.append(right[right_index])
            right_index += 1
    result.extend(left[left_index:])
    result.extend(right[right_index:])
    return tuple(result)


def _sparse_root_member_candidates(
    blocks: Sequence[_SparseMemberBlock],
    query: Rect,
) -> tuple[int, ...]:
    return tuple(
        block.member_ids[candidate]
        for block in blocks
        for candidate in block.index.ordered_intersections(query)
    )


def _sparse_root_has_witness(
    item: int,
    blocks: Sequence[_SparseMemberBlock],
    rectangles: Sequence[Rect],
    query: Rect,
    *,
    gap_tol: int,
    table_barriers: _RectangleIndex,
    corridor_cache: dict[tuple[str, int, int, int, int], bool],
) -> bool:
    seen: set[Rect] = set()
    for candidate in _sparse_root_member_candidates(
        blocks,
        _expanded_proximity_rect(query, gap_tol=gap_tol),
    ):
        candidate_rect = rectangles[candidate]
        if candidate == item or candidate_rect in seen:
            continue
        seen.add(candidate_rect)
        if _sparse_primitives_connect(
            query,
            candidate_rect,
            gap_tol=gap_tol,
            table_barriers=table_barriers,
            corridor_cache=corridor_cache,
        ):
            return True
    return False


def _expanded_proximity_rect(rect: Rect, *, gap_tol: int) -> Rect:
    distance = gap_tol + 1
    return Rect(
        max(1, rect.row_min - distance),
        min(1_048_576, rect.row_max + distance),
        max(1, rect.col_min - distance),
        min(16_384, rect.col_max + distance),
    )


def _sparse_primitives_connect(
    left: Rect,
    right: Rect,
    *,
    gap_tol: int,
    table_barriers: _RectangleIndex,
    corridor_cache: dict[tuple[str, int, int, int, int], bool],
) -> bool:
    row_gap = _axis_blank_gap(
        left.row_min,
        left.row_max,
        right.row_min,
        right.row_max,
    )
    col_gap = _axis_blank_gap(
        left.col_min,
        left.col_max,
        right.col_min,
        right.col_max,
    )
    if row_gap > gap_tol or col_gap > gap_tol:
        return False

    rows_overlap = max(left.row_min, right.row_min) <= min(
        left.row_max,
        right.row_max,
    )
    cols_overlap = max(left.col_min, right.col_min) <= min(
        left.col_max,
        right.col_max,
    )
    if rows_overlap and cols_overlap:
        return True
    if rows_overlap:
        if col_gap == 0:
            return True
        overlap_min = max(left.row_min, right.row_min)
        overlap_max = min(left.row_max, right.row_max)
        if left.col_max < right.col_min:
            gap_min = left.col_max + 1
            gap_max = right.col_min - 1
        else:
            gap_min = right.col_max + 1
            gap_max = left.col_min - 1
        return not _corridor_axis_is_covered(
            "h",
            Rect(overlap_min, overlap_max, gap_min, gap_max),
            overlap_min,
            overlap_max,
            table_barriers,
            corridor_cache,
        )
    if cols_overlap:
        if row_gap == 0:
            return True
        overlap_min = max(left.col_min, right.col_min)
        overlap_max = min(left.col_max, right.col_max)
        if left.row_max < right.row_min:
            gap_min = left.row_max + 1
            gap_max = right.row_min - 1
        else:
            gap_min = right.row_max + 1
            gap_max = left.row_min - 1
        return not _corridor_axis_is_covered(
            "v",
            Rect(gap_min, gap_max, overlap_min, overlap_max),
            overlap_min,
            overlap_max,
            table_barriers,
            corridor_cache,
        )

    connector = _diagonal_connector(left, right)
    key = ("d", *_rect_key(connector))
    blocked = corridor_cache.get(key)
    if blocked is None:
        blocked = table_barriers.intersects_any(connector)
        corridor_cache[key] = blocked
    return not blocked


def _corridor_axis_is_covered(
    axis: Literal["h", "v"],
    corridor: Rect,
    lower: int,
    upper: int,
    table_barriers: _RectangleIndex,
    cache: dict[tuple[str, int, int, int, int], bool],
) -> bool:
    key = (axis, *_rect_key(corridor))
    cached = cache.get(key)
    if cached is not None:
        return cached
    intervals: list[tuple[int, int]] = []
    for table_index in table_barriers.intersections(corridor):
        table = table_barriers.rectangle(table_index)
        if axis == "h":
            intervals.append(
                (
                    max(lower, table.row_min),
                    min(upper, table.row_max),
                )
            )
        else:
            intervals.append(
                (
                    max(lower, table.col_min),
                    min(upper, table.col_max),
                )
            )
    cursor = lower
    covered = False
    for start, end in sorted(intervals):
        if end < cursor:
            continue
        if start > cursor:
            break
        cursor = max(cursor, end + 1)
        if cursor > upper:
            covered = True
            break
    cache[key] = covered
    return covered


def _diagonal_connector(left: Rect, right: Rect) -> Rect:
    if left.row_max < right.row_min:
        row_min, row_max = left.row_max, right.row_min
    else:
        row_min, row_max = right.row_max, left.row_min
    if left.col_max < right.col_min:
        col_min, col_max = left.col_max, right.col_min
    else:
        col_min, col_max = right.col_max, left.col_min
    return Rect(row_min, row_max, col_min, col_max)


def _axis_blank_gap(
    left_min: int,
    left_max: int,
    right_min: int,
    right_max: int,
) -> int:
    if left_max < right_min:
        return right_min - left_max - 1
    if right_max < left_min:
        return left_min - right_max - 1
    return 0


def _union_near_runs(
    runs: Sequence[_Run],
    current: Sequence[int],
    prior: Sequence[int],
    disjoint: _DisjointSet,
    *,
    gap_tol: int,
) -> None:
    prior_cursor = 0
    for current_index in current:
        current_run = runs[current_index]
        while (
            prior_cursor < len(prior)
            and runs[prior[prior_cursor]].col_max + gap_tol + 1 < current_run.col_min
        ):
            prior_cursor += 1
        candidate = prior_cursor
        while (
            candidate < len(prior)
            and runs[prior[candidate]].col_min <= current_run.col_max + gap_tol + 1
        ):
            disjoint.union(current_index, prior[candidate])
            candidate += 1


def _merge_intersecting_rectangles(
    rectangles: Sequence[Rect],
) -> tuple[Rect, ...]:
    """Compute deterministic bounding-box closure without global fixed-point sweeps."""
    current = tuple(sorted(rectangles, key=_rect_key))
    if len(current) < 2:
        return current
    disjoint = _DisjointSet(len(current))
    bounds = list(current)
    index = _RectangleIndex(current)
    pending = deque(range(len(current)))

    while pending:
        item = pending.popleft()
        root = disjoint.find(item)
        if root != item:
            continue
        while True:
            targets = tuple(
                disjoint.find(candidate)
                for candidate in index.intersections(bounds[root], exclude=root)
            )
            targets = tuple(target for target in targets if target != root)
            if not targets:
                break
            index.remove(root)
            combined = bounds[root]
            for target in targets:
                index.remove(target)
                combined = _bounding_rect(combined, bounds[target])
                root = disjoint.union(root, target)
            bounds[root] = combined
            index.insert(root, combined)

    return tuple(sorted(index.rectangles(), key=_rect_key))


def _overlapping_pairs(rectangles: Sequence[Rect]) -> Iterable[tuple[int, int]]:
    """Yield exact overlaps after coordinate-ordering while preserving input indices."""
    active_buckets: dict[int, set[int]] = defaultdict(set)
    active_wide: set[int] = set()
    expiration_heap: list[tuple[int, int]] = []
    buckets_by_index: list[tuple[int, ...]] = [()] * len(rectangles)
    ordered = sorted(
        enumerate(rectangles),
        key=lambda item: (*_rect_key(item[1]), item[0]),
    )

    for current_index, current in ordered:
        while expiration_heap and expiration_heap[0][0] < current.row_min:
            _row_max, expired_index = heappop(expiration_heap)
            expired_buckets = buckets_by_index[expired_index]
            if len(expired_buckets) > _OVERLAP_WIDE_BUCKETS:
                active_wide.discard(expired_index)
                continue
            for bucket in expired_buckets:
                active_buckets[bucket].discard(expired_index)
                if not active_buckets[bucket]:
                    del active_buckets[bucket]

        current_buckets = _rect_column_buckets(current)
        candidates = set(active_wide)
        for bucket in current_buckets:
            candidates.update(active_buckets.get(bucket, ()))
        for prior_index in sorted(candidates):
            if current.intersects(rectangles[prior_index]):
                yield prior_index, current_index

        buckets_by_index[current_index] = current_buckets
        if len(current_buckets) > _OVERLAP_WIDE_BUCKETS:
            active_wide.add(current_index)
        else:
            for bucket in current_buckets:
                active_buckets[bucket].add(current_index)
        heappush(expiration_heap, (current.row_max, current_index))


def _rect_column_buckets(rect: Rect) -> tuple[int, ...]:
    first = (rect.col_min - 1) // _OVERLAP_BUCKET_WIDTH
    last = (rect.col_max - 1) // _OVERLAP_BUCKET_WIDTH
    return tuple(range(first, last + 1))


def _infer_headers(
    seeds: Sequence[_Seed],
    merges: Sequence[Rect],
    styles: StyleCatalog,
    cells: Iterable[RegionCell],
    *,
    threshold: float,
) -> tuple[_HeaderDecision, ...]:
    evidence = [_HeaderEvidence() for _seed in seeds]
    merge_index = _MergeHeaderIndex(merges)
    locator = _RegionLocator(seeds)
    for cell in cells:
        seed_index = locator.locate(cell.row, cell.col)
        if seed_index is None:
            continue
        seed = seeds[seed_index]
        if seed.kind == "table":
            continue
        max_candidate_row = seed.rect.row_min + min(3, max(0, _height(seed.rect) - 1)) - 1
        if cell.row <= max_candidate_row:
            evidence[seed_index].cells[(cell.row, cell.col)] = cell
        preview = evidence[seed_index].preview[cell.col]
        if len(preview) < HEADER_BODY_PREVIEW:
            preview.append(cell)

    decisions: list[_HeaderDecision] = []
    for index, seed in enumerate(seeds):
        if seed.table is not None:
            headers = tuple(seed.table.columns)
            normalized = deduplicate_normalized_headers(
                (header, f"Column {column_label(seed.rect.col_min + offset)}")
                for offset, header in enumerate(headers)
            )
            decisions.append(_HeaderDecision(seed.table.header_rows, 1.0, headers, normalized))
            continue
        decisions.append(
            _score_heuristic_headers(
                seed.rect,
                evidence[index],
                merge_index.view(
                    seed.rect,
                    min(3, max(0, _height(seed.rect) - 1)),
                ),
                styles,
                threshold=threshold,
            )
        )
    return tuple(decisions)


def _score_heuristic_headers(
    rect: Rect,
    evidence: _HeaderEvidence,
    merges: _HeaderMergeView,
    styles: StyleCatalog,
    *,
    threshold: float,
) -> _HeaderDecision:
    candidates = min(3, max(0, _height(rect) - 1))
    best_rows = 0
    best_score = 0.0
    best_headers: tuple[str, ...] = ()
    for header_rows in range(1, candidates + 1):
        synthesized = _synthesize_headers(rect, header_rows, evidence.cells, merges)
        width = _width(rect)
        nonblank = [item for item in synthesized if item[1]]
        text_coverage = sum(item[2] for item in synthesized) / width
        coverage = len(nonblank) / width
        normalized_nonblank = {
            normalize_header(item[0], fallback=f"Column {column_label(rect.col_min + index)}")
            for index, item in enumerate(synthesized)
            if item[1]
        }
        uniqueness = len(normalized_nonblank) / len(nonblank) if nonblank else 0.0
        contrasts = 0
        contrast_denominator = 0
        style_shifts = 0
        style_denominator = 0
        for offset, item in enumerate(synthesized):
            column = rect.col_min + offset
            body = [
                cell
                for cell in evidence.preview.get(column, ())
                if cell.row >= rect.row_min + header_rows
            ]
            nonnull_body = [
                cell for cell in body if cell.value is not None and cell.value_type != "blank"
            ]
            if item[1] and nonnull_body:
                contrast_denominator += 1
                if item[2] and any(
                    _dtype_atom(cell) not in {"str", "other"} for cell in nonnull_body
                ):
                    contrasts += 1
            if item[1]:
                style_denominator += 1
                header_signatures = {_style_signature(styles, style_idx) for style_idx in item[3]}
                body_style = _modal_style(nonnull_body)
                body_signature = _style_signature(styles, body_style)
                if any(
                    _is_highlighted(signature) and signature != body_signature
                    for signature in header_signatures
                ):
                    style_shifts += 1
        contrast = contrasts / contrast_denominator if contrast_denominator else 0.0
        style_shift = style_shifts / style_denominator if style_denominator else 0.0
        score = round(
            _HEADER_WEIGHTS[0] * text_coverage
            + _HEADER_WEIGHTS[1] * contrast
            + _HEADER_WEIGHTS[2] * uniqueness
            + _HEADER_WEIGHTS[3] * style_shift
            + _HEADER_WEIGHTS[4] * coverage,
            6,
        )
        display_headers = tuple(
            item[0] or f"Column {column_label(rect.col_min + offset)}"
            for offset, item in enumerate(synthesized)
        )
        if score > best_score:
            best_rows = header_rows
            best_score = score
            best_headers = display_headers

    if best_rows and best_score >= threshold:
        headers = best_headers
        confidence = best_score
        selected_rows = best_rows
    else:
        headers = tuple(
            f"Column {column_label(column)}" for column in range(rect.col_min, rect.col_max + 1)
        )
        confidence = round(1.0 - best_score, 6)
        selected_rows = 0
    normalized = deduplicate_normalized_headers(
        (header, f"Column {column_label(rect.col_min + offset)}")
        for offset, header in enumerate(headers)
    )
    return _HeaderDecision(selected_rows, confidence, headers, normalized)


def _synthesize_headers(
    rect: Rect,
    header_rows: int,
    cells: Mapping[tuple[int, int], RegionCell],
    merges: _HeaderMergeView,
) -> tuple[tuple[str, bool, bool, tuple[int, ...]], ...]:
    result: list[tuple[str, bool, bool, tuple[int, ...]]] = []
    for column in range(rect.col_min, rect.col_max + 1):
        pieces: list[str] = []
        all_text = True
        style_indices: list[int] = []
        seen_merges: set[tuple[int, int]] = set()
        for row in range(rect.row_min, rect.row_min + header_rows):
            anchor = merges.anchor_at(row, column)
            if anchor is not None:
                if anchor in seen_merges:
                    continue
                seen_merges.add(anchor)
                cell = cells.get(anchor)
            else:
                cell = cells.get((row, column))
            if cell is None or cell.value is None or cell.value_type == "blank":
                continue
            pieces.append(_display_value(cell.value))
            all_text = all_text and cell.value_type == "string"
            style_indices.append(cell.style_idx)
        result.append(
            (" / ".join(pieces), bool(pieces), bool(pieces) and all_text, tuple(style_indices))
        )
    return tuple(result)


def _profile_columns(
    seeds: Sequence[_Seed],
    decisions: Sequence[_HeaderDecision],
    cells: Iterable[RegionCell],
    *,
    cell_count: int,
    options: RegionOptions,
) -> tuple[tuple[ColumnProfile, ...], ...]:
    is_large = cell_count > options.large_sheet_threshold
    sample_limit = options.large_dtype_sample_limit if is_large else options.dtype_sample_limit
    sample_stride = max(2, math.ceil(cell_count / options.large_sheet_threshold)) if is_large else 1
    accumulators = [
        [
            _ColumnAccumulator(sample_limit, sample_stride, options.distinct_cap)
            for _column in range(_width(seed.rect))
        ]
        for seed in seeds
    ]
    locator = _RegionLocator(seeds)
    for cell in cells:
        seed_index = locator.locate(cell.row, cell.col)
        if seed_index is None:
            continue
        seed = seeds[seed_index]
        body_min = seed.rect.row_min + decisions[seed_index].rows
        body_max = seed.rect.row_max
        if seed.table is not None:
            body_max -= seed.table.totals_rows
        if not body_min <= cell.row <= body_max:
            continue
        accumulators[seed_index][cell.col - seed.rect.col_min].add(cell)

    result: list[tuple[ColumnProfile, ...]] = []
    for seed_index, _seed in enumerate(seeds):
        decision = decisions[seed_index]
        result.append(
            tuple(
                ColumnProfile(
                    idx=offset,
                    header=decision.headers[offset],
                    norm_header=decision.normalized_headers[offset],
                    dtype=accumulator.dtype(),
                    nonnull=accumulator.nonnull,
                    distinct_est=len(accumulator.distinct),
                )
                for offset, accumulator in enumerate(accumulators[seed_index])
            )
        )
    return tuple(result)


def _style_signature(
    styles: StyleCatalog, style_idx: int
) -> tuple[bool, tuple[str | None, str | None, str | None] | None]:
    if not 0 <= style_idx < len(styles.cell_xfs):
        return False, None
    cell_style = styles.cell_xfs[style_idx]
    bold = (
        styles.fonts[cell_style.font_id].bold
        if 0 <= cell_style.font_id < len(styles.fonts)
        else False
    )
    if not 0 <= cell_style.fill_id < len(styles.fills):
        return bold, None
    fill = styles.fills[cell_style.fill_id]
    pattern = None if fill.pattern_type is None else fill.pattern_type.casefold()
    if pattern in {None, "none", "gray125"}:
        return bold, None
    return bold, (fill.pattern_type, fill.foreground, fill.background)


def _is_highlighted(signature: tuple[bool, object | None]) -> bool:
    return signature[0] or signature[1] is not None


def _modal_style(cells: Sequence[RegionCell]) -> int:
    if not cells:
        return 0
    counts = Counter(cell.style_idx for cell in cells)
    return min(counts, key=lambda style_idx: (-counts[style_idx], style_idx))


def _dtype_atom(cell: RegionCell) -> str:
    if cell.value_type == "date" or isinstance(cell.value, (date, datetime, time)):
        return "date"
    if cell.value_type == "error":
        return "other"
    if isinstance(cell.value, bool):
        return "bool"
    if isinstance(cell.value, int):
        return "int"
    if isinstance(cell.value, float):
        return "float"
    if isinstance(cell.value, str):
        return "str"
    return "other"


def _distinct_key(cell: RegionCell) -> tuple[str, str, str]:
    return cell.value_type, type(cell.value).__name__, repr(cell.value)


def _display_value(value: CellScalar) -> str:
    if isinstance(value, (date, datetime, time)):
        return value.isoformat()
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)


def _runs_rect(runs: Sequence[_Run]) -> Rect:
    return Rect(
        min(run.row for run in runs),
        max(run.row for run in runs),
        min(run.col_min for run in runs),
        max(run.col_max for run in runs),
    )


def _bounding_rect(left: Rect, right: Rect) -> Rect:
    return Rect(
        min(left.row_min, right.row_min),
        max(left.row_max, right.row_max),
        min(left.col_min, right.col_min),
        max(left.col_max, right.col_max),
    )


def _rectangles_bounds(rectangles: Iterable[Rect]) -> Rect:
    iterator = iter(rectangles)
    try:
        bounds = next(iterator)
    except StopIteration as exc:
        raise ValueError("cannot bound an empty rectangle collection") from exc
    for rect in iterator:
        bounds = _bounding_rect(bounds, rect)
    return bounds


def _rect_key(rect: Rect) -> tuple[int, int, int, int]:
    return rect.row_min, rect.col_min, rect.row_max, rect.col_max


def _seed_sort_key(seed: _Seed) -> tuple[int, int, int, int, str, str]:
    return (
        *_rect_key(seed.rect),
        seed.kind,
        "" if seed.table is None else seed.table.name,
    )


def _width(rect: Rect) -> int:
    return rect.col_max - rect.col_min + 1


def _height(rect: Rect) -> int:
    return rect.row_max - rect.row_min + 1


def _require_disjoint(seeds: Sequence[_Seed]) -> None:
    rectangles = tuple(seed.rect for seed in seeds)
    if next(iter(_overlapping_pairs(rectangles)), None) is not None:
        raise RuntimeError("region normalization produced overlapping rectangles")


__all__ = [
    "CellStreamFactory",
    "ColumnDType",
    "ColumnProfile",
    "DetectedRegion",
    "RegionAnalysis",
    "RegionCell",
    "RegionKind",
    "RegionOptions",
    "RegionWarning",
    "analyze_sheet_regions",
]
