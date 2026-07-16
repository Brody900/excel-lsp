"""Hypothesis checks for region invariants I4-I6."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable

from hypothesis import given, settings
from hypothesis import strategies as st

from excel_lsp.core.models import Rect, SheetDescriptor, SheetParseSummary, TableInfo
from excel_lsp.core.parse.coordinates import contains
from excel_lsp.core.parse.styles import DEFAULT_STYLE_CATALOG
from excel_lsp.core.regions import (
    RegionAnalysis,
    RegionCell,
    RegionOptions,
    _coalesce_row_runs,  # pyright: ignore[reportPrivateUsage]
    _component_first_sparse_rectangles,  # pyright: ignore[reportPrivateUsage]
    _merge_intersecting_rectangles,  # pyright: ignore[reportPrivateUsage]
    _overlapping_pairs,  # pyright: ignore[reportPrivateUsage]
    _partition_runs_around_tables,  # pyright: ignore[reportPrivateUsage]
    _RectangleIndex,  # pyright: ignore[reportPrivateUsage]
    _Run,  # pyright: ignore[reportPrivateUsage]
    _runs_rect,  # pyright: ignore[reportPrivateUsage]
    _Span,  # pyright: ignore[reportPrivateUsage]
    analyze_sheet_regions,
)
from excel_lsp.core.symbols import column_symbol_id, region_symbol_id


def _summary(
    cells: tuple[RegionCell, ...],
    *,
    tables: tuple[TableInfo, ...] = (),
    merges: tuple[Rect, ...] = (),
) -> SheetParseSummary:
    return SheetParseSummary(
        descriptor=SheetDescriptor(
            0,
            "Property",
            1,
            "rId1",
            "xl/worksheets/sheet1.xml",
            "worksheet",
        ),
        part_hash="hash",
        max_row=max((cell.row for cell in cells), default=0),
        max_col=max((cell.col for cell in cells), default=0),
        cell_count=len(cells),
        merges=merges,
        tables=tables,
    )


def _cells(points: Iterable[tuple[int, int]]) -> tuple[RegionCell, ...]:
    return tuple(
        RegionCell(row, col, row * 100 + col, "number") for row, col in sorted(set(points))
    )


def _symbol_ids(analysis: RegionAnalysis) -> tuple[str, ...]:
    result: list[str] = []
    for region in analysis.regions:
        result.append(region_symbol_id("Property", region.n))
        result.extend(
            column_symbol_id("Property", region.n, column.norm_header) for column in region.columns
        )
    return tuple(result)


def _planted_rectangles(
    specs: list[tuple[int, int, int, int, int]],
) -> tuple[Rect, ...]:
    rectangles: list[Rect] = []
    for slot, row_offset, col_offset, height, width in specs:
        slot_row, slot_col = divmod(slot, 4)
        row_min = 21 + slot_row * 10 + row_offset
        col_min = 1 + slot_col * 10 + col_offset
        rectangles.append(
            Rect(
                row_min,
                row_min + height - 1,
                col_min,
                col_min + width - 1,
            )
        )
    return tuple(
        sorted(
            rectangles,
            key=lambda rect: (rect.row_min, rect.col_min, rect.row_max, rect.col_max),
        )
    )


_SEPARATED_RECTANGLES = st.lists(
    st.tuples(
        st.integers(0, 11),
        st.integers(0, 1),
        st.integers(0, 1),
        st.integers(2, 4),
        st.integers(2, 4),
    ),
    min_size=4,
    max_size=8,
    unique_by=lambda spec: spec[0],
).map(_planted_rectangles)


@settings(max_examples=80, deadline=None)
@given(
    points=st.sets(
        st.tuples(st.integers(1, 12), st.integers(1, 12)),
        min_size=1,
        max_size=50,
    ),
    gap_tol=st.integers(0, 2),
)
def test_i4_random_sparse_cells_belong_to_exactly_one_nonoverlapping_region(
    points: set[tuple[int, int]],
    gap_tol: int,
) -> None:
    cells = _cells(points)

    analysis = analyze_sheet_regions(
        _summary(cells),
        DEFAULT_STYLE_CATALOG,
        lambda: iter(cells),
        RegionOptions(gap_tol=gap_tol),
    )

    assert analysis.regions
    for cell in cells:
        assert sum(contains(region.rect, cell.row, cell.col) for region in analysis.regions) == 1
    for index, region in enumerate(analysis.regions):
        assert all(not region.rect.intersects(prior.rect) for prior in analysis.regions[:index])


@settings(max_examples=60, deadline=None)
@given(
    top=st.integers(1, 7),
    left=st.integers(1, 7),
    height=st.integers(1, 4),
    width=st.integers(1, 4),
    points=st.sets(
        st.tuples(st.integers(1, 12), st.integers(1, 12)),
        max_size=45,
    ),
)
def test_i5_listobject_range_is_exact_and_heuristics_never_intersect_it(
    top: int,
    left: int,
    height: int,
    width: int,
    points: set[tuple[int, int]],
) -> None:
    rect = Rect(top, top + height - 1, left, left + width - 1)
    table = TableInfo(
        "PropertyTable",
        "PropertyTable",
        _rect_ref(rect),
        1,
        0,
        tuple(f"Column {index}" for index in range(1, width + 1)),
    )
    cells = _cells(points)

    analysis = analyze_sheet_regions(
        _summary(cells, tables=(table,)),
        DEFAULT_STYLE_CATALOG,
        lambda: iter(cells),
    )

    table_regions = [region for region in analysis.regions if region.kind == "table"]
    assert [(region.rect, region.list_object_name) for region in table_regions] == [
        (rect, "PropertyTable")
    ]
    assert all(
        not region.rect.intersects(rect) for region in analysis.regions if region.kind == "region"
    )
    for cell in cells:
        assert sum(contains(region.rect, cell.row, cell.col) for region in analysis.regions) == 1


@settings(max_examples=80, deadline=None)
@given(
    points=st.sets(
        st.tuples(st.integers(1, 15), st.integers(1, 10)),
        min_size=1,
        max_size=55,
    ),
    chunk_size=st.integers(1, 9),
)
def test_i6_identical_content_has_identical_ordinals_across_stream_chunking(
    points: set[tuple[int, int]],
    chunk_size: int,
) -> None:
    cells = _cells(points)
    tables = (
        TableInfo("FirstTable", "FirstTable", "B2:C4", 1, 0, ("First", "Second")),
        TableInfo("SecondTable", "SecondTable", "H10:I12", 1, 0, ("Third", "Fourth")),
    )

    def chunked() -> Iterable[RegionCell]:
        for start in range(0, len(cells), chunk_size):
            yield from cells[start : start + chunk_size]

    direct = analyze_sheet_regions(
        _summary(cells, tables=tables),
        DEFAULT_STYLE_CATALOG,
        lambda: iter(cells),
    )
    batched = analyze_sheet_regions(
        _summary(cells, tables=tuple(reversed(tables))),
        DEFAULT_STYLE_CATALOG,
        chunked,
    )

    assert batched == direct
    assert _symbol_ids(batched) == _symbol_ids(direct)
    assert [region.n for region in direct.regions] == list(range(len(direct.regions)))
    assert [region.rect for region in direct.regions] == sorted(
        (region.rect for region in direct.regions),
        key=lambda rect: (rect.row_min, rect.col_min, rect.row_max, rect.col_max),
    )


@settings(max_examples=60, deadline=None)
@given(
    rectangles=_SEPARATED_RECTANGLES,
    gap_tol=st.integers(0, 2),
    chunk_size=st.integers(1, 11),
)
def test_i4_i6_planted_separated_rectangular_grids_are_stable(
    rectangles: tuple[Rect, ...],
    gap_tol: int,
    chunk_size: int,
) -> None:
    table_rectangles = (
        Rect(1, 3, 1, 2),
        Rect(1, 4, 11, 13),
    )
    expected_rectangles = tuple(
        sorted(
            (*table_rectangles, *rectangles),
            key=lambda rect: (rect.row_min, rect.col_min, rect.row_max, rect.col_max),
        )
    )
    points = {
        (row, col)
        for rect in expected_rectangles
        for row in range(rect.row_min, rect.row_max + 1)
        for col in range(rect.col_min, rect.col_max + 1)
    }
    cells = _cells(points)
    tables = tuple(
        TableInfo(
            f"PlantedTable{index + 1}",
            f"PlantedTable{index + 1}",
            _rect_ref(rect),
            1,
            0,
            tuple(f"Table {index + 1} Column {column + 1}" for column in range(_width(rect))),
        )
        for index, rect in enumerate(table_rectangles)
    )
    merges = tuple(
        Rect(rect.row_min, rect.row_min, rect.col_min, rect.col_min + 1) for rect in rectangles[:2]
    )

    def chunked() -> Iterable[RegionCell]:
        for start in range(0, len(cells), chunk_size):
            yield from cells[start : start + chunk_size]

    options = RegionOptions(gap_tol=gap_tol)
    direct = analyze_sheet_regions(
        _summary(cells, tables=tables, merges=merges),
        DEFAULT_STYLE_CATALOG,
        lambda: iter(cells),
        options,
    )
    reordered = analyze_sheet_regions(
        _summary(
            cells,
            tables=tuple(reversed(tables)),
            merges=tuple(reversed(merges)),
        ),
        DEFAULT_STYLE_CATALOG,
        chunked,
        options,
    )

    assert reordered == direct
    assert _symbol_ids(reordered) == _symbol_ids(direct)
    assert tuple(region.rect for region in direct.regions) == expected_rectangles
    assert {region.rect for region in direct.regions if region.kind == "table"} == set(
        table_rectangles
    )
    for row, col in points:
        assert sum(contains(region.rect, row, col) for region in direct.regions) == 1
    for index, region in enumerate(direct.regions):
        assert all(not region.rect.intersects(prior.rect) for prior in direct.regions[:index])


@settings(max_examples=100, deadline=None)
@given(
    specs=st.lists(
        st.tuples(
            st.integers(1, 120),
            st.integers(0, 20),
            st.integers(1, 16_000),
            st.integers(0, 384),
        ),
        max_size=25,
    ),
    wide_row=st.integers(1, 120),
    wide_height=st.integers(0, 20),
    wide_col=st.integers(1, 16_000),
)
def test_overlap_sweep_matches_brute_force_with_wide_rectangles(
    specs: list[tuple[int, int, int, int]],
    wide_row: int,
    wide_height: int,
    wide_col: int,
) -> None:
    rectangles = [
        Rect(
            wide_row,
            wide_row + wide_height,
            wide_col,
            wide_col + 320,
        ),
        *(
            Rect(
                row,
                row + height,
                col,
                min(16_384, col + width),
            )
            for row, height, col, width in specs
        ),
    ]
    expected = {
        (left_index, right_index)
        for left_index, left in enumerate(rectangles)
        for right_index, right in enumerate(rectangles[left_index + 1 :], left_index + 1)
        if left.intersects(right)
    }
    actual = {tuple(sorted(pair)) for pair in _overlapping_pairs(tuple(reversed(rectangles)))}
    remapped_actual = {
        (len(rectangles) - 1 - right, len(rectangles) - 1 - left) for left, right in actual
    }

    assert remapped_actual == expected


@settings(max_examples=80, deadline=None)
@given(
    specs=st.lists(
        st.tuples(
            st.integers(1, 20),
            st.integers(0, 5),
            st.integers(1, 20),
            st.integers(0, 5),
        ),
        max_size=15,
    )
)
def test_rectangle_bounding_closure_matches_fixed_point_reference(
    specs: list[tuple[int, int, int, int]],
) -> None:
    rectangles = tuple(
        Rect(row, row + height, col, col + width) for row, height, col, width in specs
    )

    assert _merge_intersecting_rectangles(rectangles) == _brute_bounding_closure(rectangles)


@settings(max_examples=400, deadline=None, derandomize=True)
@given(
    run_specs=st.lists(
        st.tuples(
            st.integers(1, 20),
            st.integers(1, 20),
            st.integers(0, 5),
        ),
        min_size=1,
        max_size=35,
        unique=True,
    ),
    table_slots=st.lists(
        st.integers(0, 15),
        max_size=7,
        unique=True,
    ),
)
def test_indexed_table_partition_matches_reference_across_table_order(
    run_specs: list[tuple[int, int, int]],
    table_slots: list[int],
) -> None:
    runs = tuple(
        sorted(
            (_Run(row, col, min(24, col + width)) for row, col, width in run_specs),
            key=lambda run: (run.row, run.col_min, run.col_max),
        )
    )
    tables = tuple(
        Rect(
            1 + (slot // 4) * 5,
            1 + (slot // 4) * 5 + slot % 2,
            1 + (slot % 4) * 5,
            1 + (slot % 4) * 5 + (slot // 2) % 2,
        )
        for slot in table_slots
    )
    expected = _canonical_run_zones(_reference_table_partition(runs, tables))

    for ordered_tables in (tables, tuple(reversed(tables))):
        actual = _partition_runs_around_tables(
            runs,
            ordered_tables,
            _RectangleIndex(ordered_tables),
        )
        assert _canonical_run_zones(actual) == expected


@st.composite
def _merge_table_geometry(
    draw: st.DrawFn,
) -> tuple[tuple[Rect, ...], tuple[Rect, ...]]:
    slots = draw(
        st.lists(
            st.integers(0, 8),
            min_size=1,
            max_size=7,
            unique=True,
        )
    )
    merge_count = draw(st.integers(1, len(slots)))
    rectangles: list[Rect] = []
    for slot in slots:
        row_min = 1 + (slot // 3) * 5
        col_min = 1 + (slot % 3) * 5
        height = draw(st.integers(1, 3))
        width = draw(st.integers(1, 3))
        rectangles.append(
            Rect(
                row_min,
                row_min + height - 1,
                col_min,
                col_min + width - 1,
            )
        )
    merges = tuple(sorted(rectangles[:merge_count], key=_property_rect_key))
    tables = tuple(sorted(rectangles[merge_count:], key=_property_rect_key))
    return merges, tables


@settings(max_examples=200, deadline=None, derandomize=True)
@given(
    points=st.sets(
        st.tuples(st.integers(1, 15), st.integers(1, 15)),
        max_size=55,
    ),
    geometry=_merge_table_geometry(),
    gap_tol=st.integers(0, 3),
)
def test_component_first_sparse_geometry_matches_brute_oracle(
    points: set[tuple[int, int]],
    geometry: tuple[tuple[Rect, ...], tuple[Rect, ...]],
    gap_tol: int,
) -> None:
    merges, tables = geometry
    all_points = {
        *points,
        *((merge.row_min, merge.col_min) for merge in merges),
    }
    active_points = {
        point for point in all_points if not any(contains(table, *point) for table in tables)
    }
    exact_runs = _points_to_runs(active_points)
    spans = tuple(
        _Span(
            merge.row_min,
            merge.row_max,
            merge.col_min,
            merge.col_max,
        )
        for merge in merges
    )

    for ordered_tables in (tables, tuple(reversed(tables))):
        table_barriers = _RectangleIndex(ordered_tables)
        runs = _coalesce_row_runs(
            exact_runs,
            gap_tol=gap_tol,
            table_barriers=table_barriers,
        )
        expected = _brute_component_first_rectangles(
            tuple(
                (
                    Rect(run.row, run.row, run.col_min, run.col_max),
                    f"run:{index}",
                )
                for index, run in enumerate(runs)
            )
            + tuple((span.rect, f"span:{index}") for index, span in enumerate(spans)),
            ordered_tables,
            gap_tol=gap_tol,
        )

        actual = _component_first_sparse_rectangles(
            runs,
            spans,
            ordered_tables,
            table_barriers,
            gap_tol=gap_tol,
        )

        assert actual == expected


def _brute_component_first_rectangles(
    primitives: tuple[tuple[Rect, str], ...],
    tables: tuple[Rect, ...],
    *,
    gap_tol: int,
) -> tuple[Rect, ...]:
    ordered_tables = tuple(sorted(tables, key=_property_rect_key))
    pending = list(
        _brute_closed_sparse_components(
            primitives,
            ordered_tables,
            gap_tol=gap_tol,
        )
    )
    result: list[Rect] = []
    while pending:
        component = pending.pop(0)
        bounds = _brute_bounds(tuple(rect for rect, _member in component))
        table = next(
            (candidate for candidate in ordered_tables if bounds.intersects(candidate)),
            None,
        )
        if table is None:
            result.append(bounds)
            continue
        children: tuple[list[tuple[Rect, str]], ...] = ([], [], [], [])
        for rect, member in component:
            _brute_split_primitive(rect, member, table, children)
        for child in children:
            if child:
                pending.extend(
                    _brute_closed_sparse_components(
                        tuple(child),
                        ordered_tables,
                        gap_tol=gap_tol,
                    )
                )
    return tuple(sorted(result, key=_property_rect_key))


def _brute_closed_sparse_components(
    primitives: tuple[tuple[Rect, str], ...],
    tables: tuple[Rect, ...],
    *,
    gap_tol: int,
) -> tuple[tuple[tuple[Rect, str], ...], ...]:
    ordered = tuple(
        sorted(
            primitives,
            key=lambda item: (*_property_rect_key(item[0]), item[1]),
        )
    )
    parents = list(range(len(ordered)))
    for left_index, (left, _left_member) in enumerate(ordered):
        for right_index in range(left_index + 1, len(ordered)):
            right = ordered[right_index][0]
            if not _brute_has_table_free_witness(
                left,
                right,
                tables,
                gap_tol=gap_tol,
            ):
                continue
            left_root = _find_parent(parents, left_index)
            right_root = _find_parent(parents, right_index)
            if left_root != right_root:
                parents[right_root] = left_root
    grouped: dict[int, list[tuple[Rect, str]]] = defaultdict(list)
    for index, primitive in enumerate(ordered):
        grouped[_find_parent(parents, index)].append(primitive)
    components = [
        tuple(
            sorted(
                component,
                key=lambda item: (*_property_rect_key(item[0]), item[1]),
            )
        )
        for component in grouped.values()
    ]

    changed = True
    while changed:
        changed = False
        components.sort(
            key=lambda component: (
                *_property_rect_key(_brute_bounds(tuple(rect for rect, _member in component))),
                tuple((*_property_rect_key(rect), member) for rect, member in component),
            )
        )
        for left_index, left in enumerate(components):
            left_bounds = _brute_bounds(tuple(rect for rect, _member in left))
            for right_index in range(left_index + 1, len(components)):
                right = components[right_index]
                right_bounds = _brute_bounds(tuple(rect for rect, _member in right))
                if not left_bounds.intersects(right_bounds):
                    continue
                components[left_index] = tuple(
                    sorted(
                        (*left, *right),
                        key=lambda item: (
                            *_property_rect_key(item[0]),
                            item[1],
                        ),
                    )
                )
                del components[right_index]
                changed = True
                break
            if changed:
                break
    return tuple(components)


def _brute_has_table_free_witness(
    left: Rect,
    right: Rect,
    tables: tuple[Rect, ...],
    *,
    gap_tol: int,
) -> bool:
    if (
        _brute_axis_blank_gap(
            left.row_min,
            left.row_max,
            right.row_min,
            right.row_max,
        )
        > gap_tol
    ):
        return False
    if (
        _brute_axis_blank_gap(
            left.col_min,
            left.col_max,
            right.col_min,
            right.col_max,
        )
        > gap_tol
    ):
        return False
    for left_row in range(left.row_min, left.row_max + 1):
        for left_col in range(left.col_min, left.col_max + 1):
            for right_row in range(right.row_min, right.row_max + 1):
                for right_col in range(right.col_min, right.col_max + 1):
                    witness = Rect(
                        min(left_row, right_row),
                        max(left_row, right_row),
                        min(left_col, right_col),
                        max(left_col, right_col),
                    )
                    if all(not witness.intersects(table) for table in tables):
                        return True
    return False


def _brute_split_primitive(
    rect: Rect,
    member: str,
    table: Rect,
    groups: tuple[
        list[tuple[Rect, str]],
        list[tuple[Rect, str]],
        list[tuple[Rect, str]],
        list[tuple[Rect, str]],
    ],
) -> None:
    above, below, left, right = groups
    if rect.row_min < table.row_min:
        above.append(
            (
                Rect(
                    rect.row_min,
                    min(rect.row_max, table.row_min - 1),
                    rect.col_min,
                    rect.col_max,
                ),
                member,
            )
        )
    if rect.row_max > table.row_max:
        below.append(
            (
                Rect(
                    max(rect.row_min, table.row_max + 1),
                    rect.row_max,
                    rect.col_min,
                    rect.col_max,
                ),
                member,
            )
        )
    middle_min = max(rect.row_min, table.row_min)
    middle_max = min(rect.row_max, table.row_max)
    if middle_min > middle_max:
        return
    if rect.col_min < table.col_min:
        left.append(
            (
                Rect(
                    middle_min,
                    middle_max,
                    rect.col_min,
                    min(rect.col_max, table.col_min - 1),
                ),
                member,
            )
        )
    if rect.col_max > table.col_max:
        right.append(
            (
                Rect(
                    middle_min,
                    middle_max,
                    max(rect.col_min, table.col_max + 1),
                    rect.col_max,
                ),
                member,
            )
        )


def _brute_bounds(rectangles: tuple[Rect, ...]) -> Rect:
    return Rect(
        min(rect.row_min for rect in rectangles),
        max(rect.row_max for rect in rectangles),
        min(rect.col_min for rect in rectangles),
        max(rect.col_max for rect in rectangles),
    )


def _brute_axis_blank_gap(
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


def _reference_table_partition(
    runs: tuple[_Run, ...],
    tables: tuple[Rect, ...],
) -> tuple[tuple[_Run, ...], ...]:
    zones: list[tuple[_Run, ...]] = [runs]
    for table in sorted(
        tables,
        key=lambda rect: (rect.row_min, rect.col_min, rect.row_max, rect.col_max),
    ):
        next_zones: list[tuple[_Run, ...]] = []
        for zone in zones:
            if not _runs_rect(zone).intersects(table):
                next_zones.append(zone)
                continue
            above: list[_Run] = []
            below: list[_Run] = []
            left: list[_Run] = []
            right: list[_Run] = []
            for run in zone:
                if run.row < table.row_min:
                    above.append(run)
                elif run.row > table.row_max:
                    below.append(run)
                else:
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
            next_zones.extend(tuple(group) for group in (above, below, left, right) if group)
        zones = next_zones
    return tuple(zones)


def _canonical_run_zones(
    zones: tuple[tuple[_Run, ...], ...],
) -> tuple[tuple[_Run, ...], ...]:
    return tuple(
        sorted(
            zones,
            key=lambda zone: tuple((run.row, run.col_min, run.col_max) for run in zone),
        )
    )


def _brute_bounding_closure(rectangles: tuple[Rect, ...]) -> tuple[Rect, ...]:
    current = tuple(
        sorted(
            rectangles,
            key=lambda rect: (rect.row_min, rect.col_min, rect.row_max, rect.col_max),
        )
    )
    while True:
        parents = list(range(len(current)))

        merged = False
        for left_index, left in enumerate(current):
            for right_index in range(left_index + 1, len(current)):
                if not left.intersects(current[right_index]):
                    continue
                left_root = _find_parent(parents, left_index)
                right_root = _find_parent(parents, right_index)
                if left_root != right_root:
                    parents[right_root] = left_root
                merged = True
        if not merged:
            return current
        grouped: dict[int, list[Rect]] = defaultdict(list)
        for index, rect in enumerate(current):
            grouped[_find_parent(parents, index)].append(rect)
        current = tuple(
            sorted(
                (
                    Rect(
                        min(rect.row_min for rect in group),
                        max(rect.row_max for rect in group),
                        min(rect.col_min for rect in group),
                        max(rect.col_max for rect in group),
                    )
                    for group in grouped.values()
                ),
                key=lambda rect: (rect.row_min, rect.col_min, rect.row_max, rect.col_max),
            )
        )


def _find_parent(parents: list[int], item: int) -> int:
    while parents[item] != item:
        parents[item] = parents[parents[item]]
        item = parents[item]
    return item


def _points_to_runs(points: set[tuple[int, int]]) -> tuple[_Run, ...]:
    runs: list[_Run] = []
    current: _Run | None = None
    for row, col in sorted(points):
        if current is not None and current.row == row and current.col_max + 1 == col:
            current = _Run(row, current.col_min, col)
            continue
        if current is not None:
            runs.append(current)
        current = _Run(row, col, col)
    if current is not None:
        runs.append(current)
    return tuple(runs)


def _property_rect_key(rect: Rect) -> tuple[int, int, int, int]:
    return rect.row_min, rect.col_min, rect.row_max, rect.col_max


def _rect_ref(rect: Rect) -> str:
    from excel_lsp.core.parse.coordinates import make_cell_ref

    start = make_cell_ref(rect.row_min, rect.col_min)
    end = make_cell_ref(rect.row_max, rect.col_max)
    return start if start == end else f"{start}:{end}"


def _width(rect: Rect) -> int:
    return rect.col_max - rect.col_min + 1
