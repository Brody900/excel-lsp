"""Focused tests for sparse regions, headers, and profiles."""

from __future__ import annotations

from collections.abc import Iterable, Iterator, Sequence
from datetime import date
from types import MappingProxyType
from typing import cast

import pytest

import excel_lsp.core.regions as region_module
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.models import Rect, SheetDescriptor, SheetParseSummary, TableInfo
from excel_lsp.core.parse.styles import (
    CellStyle,
    FillStyle,
    FontStyle,
    StyleCatalog,
)
from excel_lsp.core.regions import (
    RegionCell,
    RegionOptions,
    _component_rectangles,  # pyright: ignore[reportPrivateUsage]
    _merge_intersecting_rectangles,  # pyright: ignore[reportPrivateUsage]
    _partition_runs_around_tables,  # pyright: ignore[reportPrivateUsage]
    _RectangleIndex,  # pyright: ignore[reportPrivateUsage]
    _Run,  # pyright: ignore[reportPrivateUsage]
    _runs_rect,  # pyright: ignore[reportPrivateUsage]
    _RunSequence,  # pyright: ignore[reportPrivateUsage]
    analyze_sheet_regions,
)


class _CountingMerges(Sequence[Rect]):
    def __init__(self, merges: Sequence[Rect]) -> None:
        self._merges = tuple(merges)
        self.visits = 0

    def __len__(self) -> int:
        return len(self._merges)

    def __getitem__(self, index: int) -> Rect:
        return self._merges[index]

    def __iter__(self) -> Iterator[Rect]:
        for merge in self._merges:
            self.visits += 1
            yield merge


def _summary(
    cells: tuple[RegionCell, ...],
    *,
    tables: tuple[TableInfo, ...] = (),
    merges: tuple[Rect, ...] = (),
) -> SheetParseSummary:
    return SheetParseSummary(
        descriptor=SheetDescriptor(
            order=0,
            name="Data",
            sheet_id=1,
            rel_id="rId1",
            xml_part="xl/worksheets/sheet1.xml",
            kind="worksheet",
        ),
        part_hash="sheet-hash",
        max_row=max((cell.row for cell in cells), default=0),
        max_col=max((cell.col for cell in cells), default=0),
        cell_count=len(cells),
        merges=merges,
        tables=tables,
    )


def _analyze(
    cells: list[RegionCell] | tuple[RegionCell, ...],
    *,
    tables: tuple[TableInfo, ...] = (),
    merges: tuple[Rect, ...] = (),
    options: RegionOptions | None = None,
    styles: StyleCatalog | None = None,
):
    ordered = tuple(sorted(cells, key=lambda cell: (cell.row, cell.col)))
    active_styles = styles or _styles()
    return analyze_sheet_regions(
        _summary(ordered, tables=tables, merges=merges),
        active_styles,
        lambda: iter(ordered),
        options,
    )


def _styles() -> StyleCatalog:
    return StyleCatalog(
        cell_xfs=(
            CellStyle(0, 0, 0, False),
            CellStyle(0, 1, 1, False),
        ),
        fonts=(FontStyle(), FontStyle(bold=True)),
        fills=(FillStyle(), FillStyle(pattern_type="solid", foreground="rgb:FFFF00")),
        custom_num_formats=MappingProxyType({}),
    )


def _value_cell(row: int, col: int, value: object, *, style_idx: int = 0) -> RegionCell:
    if isinstance(value, bool):
        value_type = "bool"
    elif isinstance(value, (int, float)):
        value_type = "number"
    elif isinstance(value, date):
        value_type = "date"
    else:
        value_type = "string"
    return RegionCell(row, col, value, value_type, style_idx=style_idx)  # type: ignore[arg-type]


def test_dense_origin_certificate_skips_coordinate_redetection() -> None:
    cells = tuple(
        _value_cell(row, col, f"{row}:{col}") for row in range(1, 4) for col in range(1, 3)
    )

    def unexpected_coordinate_scan() -> Iterable[RegionCell]:
        raise AssertionError("dense actual bounds must not be rescanned")

    analysis = analyze_sheet_regions(
        _summary(cells),
        _styles(),
        lambda: iter(cells),
        coordinate_stream_factory=unexpected_coordinate_scan,
    )

    assert tuple(region.rect for region in analysis.regions) == (Rect(1, 3, 1, 2),)


def test_sparse_sheet_still_uses_coordinate_detector() -> None:
    cells = (
        _value_cell(1, 1, "A"),
        _value_cell(1, 2, "B"),
        _value_cell(2, 1, 1),
    )
    scans = 0

    def coordinate_scan() -> Iterable[RegionCell]:
        nonlocal scans
        scans += 1
        return iter(cells)

    analyze_sheet_regions(
        _summary(cells),
        _styles(),
        lambda: iter(cells),
        coordinate_stream_factory=coordinate_scan,
    )

    assert scans == 1


def test_gap_tolerance_merges_one_blank_row_and_column_but_not_two() -> None:
    cells = [
        _value_cell(1, 1, 1),
        _value_cell(1, 3, 2),
        _value_cell(3, 1, 3),
        _value_cell(3, 3, 4),
        _value_cell(6, 6, 5),
    ]

    analysis = _analyze(cells)

    assert [(region.n, region.rect) for region in analysis.regions] == [
        (0, Rect(1, 3, 1, 3)),
        (1, Rect(6, 6, 6, 6)),
    ]


def test_gap_tolerance_has_a_small_authoritative_maximum() -> None:
    assert RegionOptions(gap_tol=8).gap_tol == 8

    with pytest.raises(ValueError, match="gap_tol must not exceed 8"):
        RegionOptions(gap_tol=9)


@pytest.mark.parametrize("gap_tol", (True, False, 1.5, 0.0, "1", None))
def test_gap_tolerance_requires_an_integer(gap_tol: object) -> None:
    with pytest.raises(ValueError, match="gap_tol must be an integer"):
        RegionOptions(gap_tol=cast(int, gap_tol))


def test_maximum_gap_tolerance_keeps_active_row_work_linear(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_union = region_module._union_near_runs
    row_pair_work = 0

    def counted_union(
        runs: Sequence[_Run],
        current: Sequence[int],
        prior: Sequence[int],
        disjoint: region_module._DisjointSet,
        *,
        gap_tol: int,
    ) -> None:
        nonlocal row_pair_work
        row_pair_work += 1
        original_union(
            runs,
            current,
            prior,
            disjoint,
            gap_tol=gap_tol,
        )

    monkeypatch.setattr(region_module, "_union_near_runs", counted_union)

    def component_work(run_count: int) -> int:
        nonlocal row_pair_work
        start = row_pair_work
        result = _component_rectangles(
            tuple(_Run(row, 1, 1) for row in range(1, run_count + 1)),
            gap_tol=8,
        )
        work = row_pair_work - start

        assert result == (Rect(1, run_count, 1, 1),)
        return work

    small = component_work(500)
    large = component_work(1_000)

    assert small == 4_455
    assert large == 8_955
    assert large <= small * 2.02


def test_intersecting_component_bounds_merge_into_one_disjoint_region() -> None:
    cells = [
        *(_value_cell(1, col, "top") for col in range(1, 6)),
        *(_value_cell(row, 1, "left") for row in range(2, 6)),
        _value_cell(5, 5, "inside bounds"),
    ]

    analysis = _analyze(cells, options=RegionOptions(gap_tol=0))

    assert [(region.n, region.rect) for region in analysis.regions] == [(0, Rect(1, 5, 1, 5))]


def test_isolated_grid_overlap_candidate_work_is_subquadratic(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    cells = [
        _value_cell(1 + row * 3, 1 + col * 3, row * 50 + col)
        for row in range(50)
        for col in range(50)
    ]
    original_intersects = Rect.intersects
    original_sparse_intersects = region_module._rectangles_intersect
    original_interval_overlaps = region_module._AxisIntervalIndex.overlapping
    rect_intersect_calls = 0
    sparse_intersect_calls = 0
    interval_candidates = 0

    def counted_intersects(left: Rect, right: Rect) -> bool:
        nonlocal rect_intersect_calls
        rect_intersect_calls += 1
        return original_intersects(left, right)

    def counted_sparse_intersects(left: Rect, right: Rect) -> bool:
        nonlocal sparse_intersect_calls
        sparse_intersect_calls += 1
        return original_sparse_intersects(left, right)

    def counted_interval_overlaps(
        index: region_module._AxisIntervalIndex,
        lower: int,
        upper: int,
    ) -> tuple[int, ...]:
        nonlocal interval_candidates
        result = original_interval_overlaps(index, lower, upper)
        interval_candidates += len(result)
        return result

    monkeypatch.setattr(Rect, "intersects", counted_intersects)
    monkeypatch.setattr(region_module, "_rectangles_intersect", counted_sparse_intersects)
    monkeypatch.setattr(
        region_module._AxisIntervalIndex,
        "overlapping",
        counted_interval_overlaps,
    )

    analysis = _analyze(cells)
    candidate_work = rect_intersect_calls + sparse_intersect_calls + interval_candidates

    assert len(analysis.regions) == 2_500
    assert candidate_work < len(cells) * 32


def test_formula_without_a_cache_is_structural_but_not_nonnull() -> None:
    formula = RegionCell(10, 4, None, "blank", formula="=A1")

    analysis = _analyze([formula])

    assert analysis.regions[0].rect == Rect(10, 10, 4, 4)
    assert analysis.regions[0].columns[0].nonnull == 0
    assert analysis.regions[0].columns[0].dtype == "empty"


def test_listobject_is_exact_authoritative_and_excludes_totals_from_profiles() -> None:
    table = TableInfo(
        name="SalesTable",
        display_name="SalesTable",
        ref="B2:C5",
        header_rows=1,
        totals_rows=1,
        columns=("Declared Item", "Declared Amount"),
    )
    cells = [
        _value_cell(1, 1, "outside"),
        _value_cell(2, 2, "wrong header"),
        _value_cell(2, 3, "also wrong"),
        _value_cell(3, 2, "A"),
        _value_cell(3, 3, 10),
        _value_cell(4, 2, "B"),
        _value_cell(4, 3, 20),
        _value_cell(5, 2, "Total"),
        RegionCell(5, 3, 30, "number", formula="=SUM(C3:C4)"),
    ]

    analysis = _analyze(cells, tables=(table,))
    table_region = next(region for region in analysis.regions if region.kind == "table")

    assert table_region.rect == Rect(2, 5, 2, 3)
    assert table_region.list_object_name == "SalesTable"
    assert table_region.header_rows == 1
    assert table_region.confidence == 1.0
    assert [column.header for column in table_region.columns] == [
        "Declared Item",
        "Declared Amount",
    ]
    assert [(column.dtype, column.nonnull) for column in table_region.columns] == [
        ("str", 2),
        ("int", 2),
    ]
    assert all(
        not heuristic.rect.intersects(table_region.rect)
        for heuristic in analysis.regions
        if heuristic.kind == "region"
    )


def test_table_is_a_hard_barrier_for_a_surrounding_heuristic_shape() -> None:
    table = TableInfo("T", "T", "B2:B3", 1, 0, ("Inside",))
    cells = [
        _value_cell(row, col, "x")
        for row, col in (
            (1, 1),
            (1, 2),
            (1, 3),
            (2, 1),
            (2, 2),
            (2, 3),
            (3, 1),
            (3, 2),
            (3, 3),
        )
    ]

    analysis = _analyze(cells, tables=(table,))

    table_region = next(region for region in analysis.regions if region.kind == "table")
    assert table_region.rect == Rect(2, 3, 2, 2)
    for index, region in enumerate(analysis.regions):
        assert all(not region.rect.intersects(other.rect) for other in analysis.regions[:index])


def test_many_listobject_barriers_have_near_linear_candidate_work(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_intersects = region_module._rectangles_intersect
    checks = 0

    def counted_intersects(left: Rect, right: Rect) -> bool:
        nonlocal checks
        checks += 1
        return original_intersects(left, right)

    monkeypatch.setattr(region_module, "_rectangles_intersect", counted_intersects)

    def partition_work(table_count: int) -> int:
        nonlocal checks
        tables = tuple(Rect(row, row, 2, 2) for row in range(2, table_count * 2 + 1, 2))
        runs = tuple(_Run(row, 1, 3) for row in range(1, table_count * 2 + 2))
        start = checks
        table_index = _RectangleIndex(tables)
        zones = _partition_runs_around_tables(runs, tables, table_index)
        work = checks - start

        assert len(zones) == table_count * 3 + 1
        assert all(not table_index.intersects_any(_runs_rect(zone)) for zone in zones)
        return work

    small = partition_work(256)
    large = partition_work(512)

    assert large <= small * 2.5
    assert large < 512 * 40


def test_spatial_barrier_lookup_prunes_full_height_left_right_decoys(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_intersects = region_module._rectangles_intersect
    checks = 0

    def counted_intersects(left: Rect, right: Rect) -> bool:
        nonlocal checks
        checks += 1
        return original_intersects(left, right)

    monkeypatch.setattr(region_module, "_rectangles_intersect", counted_intersects)

    def partition_work(barrier_count: int) -> int:
        nonlocal checks
        center = 8_192
        row_max = barrier_count * 2 + 1
        left_count = barrier_count // 2
        decoys = (
            *(Rect(1, row_max, 1 + index, 1 + index) for index in range(left_count)),
            *(
                Rect(1, row_max, 15_500 + index, 15_500 + index)
                for index in range(barrier_count - left_count)
            ),
        )
        relevant = tuple(Rect(row, row, center, center) for row in range(2, row_max, 2))
        tables = (*decoys, *relevant)
        runs = tuple(_Run(row, center - 1, center + 1) for row in range(1, row_max + 1))
        start = checks
        zones = _partition_runs_around_tables(
            runs,
            tables,
            _RectangleIndex(tables),
        )
        work = checks - start

        assert len(zones) == barrier_count * 3 + 1
        return work

    small = partition_work(200)
    large = partition_work(400)

    assert large <= small * 2.5
    assert large < 800 * 32


def test_full_height_listobject_batch_has_linear_run_copy_work(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_init = _RunSequence.__init__
    copied_runs = 0

    def counted_init(sequence: _RunSequence, runs: Sequence[_Run]) -> None:
        nonlocal copied_runs
        copied_runs += len(runs)
        original_init(sequence, runs)

    monkeypatch.setattr(_RunSequence, "__init__", counted_init)

    def partition_work(table_count: int) -> int:
        nonlocal copied_runs
        tables = tuple(Rect(1, 12, col, col) for col in range(2, table_count * 2 + 1, 2))
        runs = tuple(
            _Run(row, col, col) for row in range(1, 13) for col in range(1, table_count * 2 + 2, 2)
        )
        start = copied_runs
        zones = _partition_runs_around_tables(
            runs,
            tables,
            _RectangleIndex(tables),
        )
        work = copied_runs - start

        assert len(zones) == table_count + 1
        assert all(len(zone) == 12 for zone in zones)
        return work

    small = partition_work(500)
    large = partition_work(1_000)

    assert small == 12_024
    assert large == 24_024
    assert large <= small * 2.01


def test_full_height_batch_stops_before_earlier_partial_height_barrier() -> None:
    tables = (
        Rect(15, 16, 5, 10),
        Rect(15, 16, 12, 12),
        Rect(14, 15, 11, 11),
        Rect(9, 9, 15, 15),
        Rect(14, 16, 1, 4),
        Rect(13, 13, 16, 16),
        Rect(11, 11, 15, 16),
    )
    runs = (
        _Run(15, 11, 13),
        _Run(15, 16, 16),
        _Run(16, 3, 3),
        _Run(16, 13, 16),
    )
    expected = (
        (_Run(15, 13, 13), _Run(15, 16, 16)),
        (_Run(16, 13, 16),),
    )

    for ordered_tables in (tables, tuple(reversed(tables))):
        zones = _partition_runs_around_tables(
            runs,
            ordered_tables,
            _RectangleIndex(ordered_tables),
        )
        canonical = tuple(
            sorted(
                zones,
                key=lambda zone: tuple((run.row, run.col_min, run.col_max) for run in zone),
            )
        )

        assert canonical == expected


@pytest.mark.parametrize(
    "tables",
    (
        (
            TableInfo("T1", "T1", "A1:B2", 1, 0, ("A", "B")),
            TableInfo("T2", "T2", "B2:C3", 1, 0, ("B", "C")),
        ),
        (TableInfo("BadWidth", "BadWidth", "A1:B2", 1, 0, ("A",)),),
        (TableInfo("BadRows", "BadRows", "A1:A1", 1, 1, ("A",)),),
    ),
)
def test_invalid_listobjects_are_structured_corruption(
    tables: tuple[TableInfo, ...],
) -> None:
    with pytest.raises(ExcelLSPError) as captured:
        _analyze([], tables=tables)

    assert captured.value.code is ErrorCode.CORRUPT


def test_malformed_listobject_ref_is_structured_corruption() -> None:
    table = TableInfo(
        "BadRef",
        "BadRef",
        "not-a-range",
        1,
        0,
        ("Column",),
    )

    with pytest.raises(ExcelLSPError) as captured:
        _analyze([], tables=(table,))

    assert captured.value.code is ErrorCode.CORRUPT
    assert captured.value.details == {"table": "BadRef", "ref": "not-a-range"}


def test_two_row_merged_headers_are_synthesized_and_win_the_score() -> None:
    cells = [
        _value_cell(1, 1, "Revenue", style_idx=1),
        _value_cell(2, 1, "Q1", style_idx=1),
        _value_cell(2, 2, "Q2", style_idx=1),
        _value_cell(3, 1, 10),
        _value_cell(3, 2, 20),
        _value_cell(4, 1, 11),
        _value_cell(4, 2, 21),
    ]

    analysis = _analyze(cells, merges=(Rect(1, 1, 1, 2),))
    region = analysis.regions[0]

    assert region.header_rows == 2
    assert [column.header for column in region.columns] == [
        "Revenue / Q1",
        "Revenue / Q2",
    ]
    assert region.confidence >= 0.8


def test_max_height_anchored_merge_is_one_lazy_span_and_one_region(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    merge = Rect(1, 1_048_576, 1, 1)
    original_components = region_module._closed_sparse_components
    primitive_counts: list[tuple[int, int]] = []

    def captured_components(
        runs: Sequence[_Run],
        spans: Sequence[region_module._Span],
        table_barriers: region_module._RectangleIndex,
        *,
        gap_tol: int,
    ) -> tuple[region_module._SparseComponent, ...]:
        primitive_counts.append((len(runs), len(spans)))
        return original_components(
            runs,
            spans,
            table_barriers,
            gap_tol=gap_tol,
        )

    monkeypatch.setattr(
        region_module,
        "_closed_sparse_components",
        captured_components,
    )

    analysis = _analyze(
        [_value_cell(1, 1, "anchor")],
        merges=(merge,),
    )

    assert primitive_counts == [(1, 1)]
    assert [(region.kind, region.rect) for region in analysis.regions] == [("region", merge)]


def test_unrelated_tall_merge_components_have_near_linear_inbounds_table_work(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_intersects = region_module._rectangles_intersect
    original_split_span = region_module._split_span_around_table
    rectangle_checks = 0
    split_span_visits = 0

    def counted_intersects(left: Rect, right: Rect) -> bool:
        nonlocal rectangle_checks
        rectangle_checks += 1
        return original_intersects(left, right)

    def counted_split_span(
        span: region_module._Span,
        table: Rect,
        groups: tuple[
            list[region_module._Span],
            list[region_module._Span],
            list[region_module._Span],
            list[region_module._Span],
        ],
    ) -> None:
        nonlocal split_span_visits
        split_span_visits += 1
        original_split_span(span, table, groups)

    monkeypatch.setattr(region_module, "_rectangles_intersect", counted_intersects)
    monkeypatch.setattr(region_module, "_split_span_around_table", counted_split_span)

    def scaling_probe(count: int) -> tuple[int, int, int]:
        nonlocal rectangle_checks, split_span_visits
        merges = tuple(Rect(1, 100_000, 1 + index * 2, 1 + index * 2) for index in range(count))
        tables = tuple(
            TableInfo(
                f"EventTable{index}",
                f"EventTable{index}",
                f"SF{2 + index * 2}",
                1,
                0,
                ("Event",),
            )
            for index in range(count)
        )
        cells = [
            *(
                _value_cell(1, merge.col_min, f"merge {index}")
                for index, merge in enumerate(merges)
            ),
            *(_value_cell(2 + index * 2, 501, f"event {index}") for index in range(count)),
        ]
        check_start = rectangle_checks
        split_start = split_span_visits

        analysis = _analyze(
            cells,
            tables=tables,
            merges=merges,
            options=RegionOptions(gap_tol=0),
        )

        return (
            rectangle_checks - check_start,
            split_span_visits - split_start,
            len(analysis.regions),
        )

    small = scaling_probe(64)
    large = scaling_probe(128)

    assert small[1:] == (0, 192)
    assert large[1:] == (0, 384)
    assert small[0] > 0
    assert large[0] <= small[0] * 2.2


def test_unanchored_merge_does_not_create_a_region() -> None:
    merge = Rect(1, 1_048_576, 1, 1)

    empty_analysis = _analyze([], merges=(merge,))
    outside_analysis = _analyze(
        [_value_cell(10, 10, "outside")],
        merges=(merge,),
    )

    assert empty_analysis.regions == ()
    assert [(region.kind, region.rect) for region in outside_analysis.regions] == [
        ("region", Rect(10, 10, 10, 10))
    ]


@pytest.mark.parametrize(
    ("right_cells", "expected_regions"),
    (
        (
            ((1, 4),),
            (Rect(1, 1, 2, 4), Rect(2, 3, 2, 2)),
        ),
        (
            ((2, 4),),
            (Rect(1, 3, 2, 2), Rect(2, 2, 4, 4)),
        ),
        (
            ((1, 4), (2, 4)),
            (
                Rect(1, 1, 2, 4),
                Rect(2, 3, 2, 2),
                Rect(2, 2, 4, 4),
            ),
        ),
    ),
)
def test_component_first_table_bsp_preserves_directional_siblings(
    right_cells: tuple[tuple[int, int], ...],
    expected_regions: tuple[Rect, ...],
) -> None:
    table = TableInfo("Barrier", "Barrier", "C2:C3", 1, 0, ("Inside",))
    merge = Rect(1, 3, 2, 2)

    analysis = _analyze(
        [
            _value_cell(1, 2, "merge anchor"),
            _value_cell(2, 3, "table"),
            *(_value_cell(row, col, "outside") for row, col in right_cells),
        ],
        tables=(table,),
        merges=(merge,),
    )

    assert (
        tuple(region.rect for region in analysis.regions if region.kind == "region")
        == expected_regions
    )


def test_unrelated_table_does_not_fragment_an_atomic_merge_component() -> None:
    merge = Rect(1, 20, 1, 1)
    table = TableInfo("Far", "Far", "M2", 1, 0, ("Inside",))

    analysis = _analyze(
        [
            _value_cell(1, 1, "merge anchor"),
            _value_cell(2, 13, "table"),
            _value_cell(2, 14, "outside"),
        ],
        tables=(table,),
        merges=(merge,),
        options=RegionOptions(gap_tol=0),
    )

    assert [(region.kind, region.rect) for region in analysis.regions] == [
        ("region", merge),
        ("table", Rect(2, 2, 13, 13)),
        ("region", Rect(2, 2, 14, 14)),
    ]


@pytest.mark.parametrize(
    ("left", "right", "gap_tol", "blocker"),
    (
        (Rect(1, 1, 1, 1), Rect(2, 2, 2, 2), 0, Rect(2, 2, 1, 1)),
        (Rect(1, 1, 1, 1), Rect(3, 3, 3, 3), 1, Rect(2, 2, 2, 2)),
        (Rect(1, 1, 1, 1), Rect(2, 2, 3, 3), 1, Rect(1, 1, 2, 2)),
        (Rect(1, 1, 1, 1), Rect(3, 3, 2, 2), 1, Rect(2, 2, 1, 1)),
    ),
)
def test_diagonal_witness_bbox_respects_tables(
    left: Rect,
    right: Rect,
    gap_tol: int,
    blocker: Rect,
) -> None:
    runs = (
        _Run(left.row_min, left.col_min, left.col_max),
        _Run(right.row_min, right.col_min, right.col_max),
    )

    clear = region_module._closed_sparse_components(
        runs,
        (),
        _RectangleIndex(()),
        gap_tol=gap_tol,
    )
    blocked = region_module._closed_sparse_components(
        runs,
        (),
        _RectangleIndex((blocker,)),
        gap_tol=gap_tol,
    )

    assert [component.bounds for component in clear] == [
        Rect(
            min(left.row_min, right.row_min),
            max(left.row_max, right.row_max),
            min(left.col_min, right.col_min),
            max(left.col_max, right.col_max),
        )
    ]
    assert [component.bounds for component in blocked] == sorted(
        (left, right),
        key=lambda rect: (rect.row_min, rect.col_min, rect.row_max, rect.col_max),
    )


def test_sparse_bounding_closure_retains_every_primitive_member() -> None:
    top = _Run(1, 1, 5)
    closure_cell = _Run(5, 5, 5)
    left = region_module._Span(1, 5, 1, 1)

    components = region_module._closed_sparse_components(
        (top, closure_cell),
        (left,),
        _RectangleIndex(()),
        gap_tol=0,
    )

    assert components == (
        region_module._SparseComponent(
            (top, closure_cell),
            (left,),
            Rect(1, 5, 1, 5),
        ),
    )


def test_bounding_closure_can_resurrect_an_expired_root() -> None:
    top_right = _Run(1, 3, 3)
    bottom = _Run(3, 1, 3)
    left = region_module._Span(1, 3, 1, 1)

    components = region_module._closed_sparse_components(
        (top_right, bottom),
        (left,),
        _RectangleIndex(()),
        gap_tol=0,
    )

    assert components == (
        region_module._SparseComponent(
            (top_right, bottom),
            (left,),
            Rect(1, 3, 1, 3),
        ),
    )


def test_mixed_full_height_table_batch_scales_linearly(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_init = region_module._RunSequence.__init__
    original_span = region_module._Span
    copied_runs = 0
    constructed_spans = 0

    def counted_init(
        sequence: region_module._RunSequence,
        runs: Sequence[_Run],
    ) -> None:
        nonlocal copied_runs
        copied_runs += len(runs)
        original_init(sequence, runs)

    def counted_span(
        row_min: int,
        row_max: int,
        col_min: int,
        col_max: int,
    ) -> region_module._Span:
        nonlocal constructed_spans
        constructed_spans += 1
        return original_span(row_min, row_max, col_min, col_max)

    monkeypatch.setattr(region_module._RunSequence, "__init__", counted_init)
    monkeypatch.setattr(region_module, "_Span", counted_span)

    def batch_work(table_count: int) -> tuple[int, int]:
        nonlocal copied_runs, constructed_spans
        runs = tuple(_Run(1, column, column) for column in range(1, table_count * 2 + 2, 2))
        spans = tuple(
            original_span(1, 12, column, column) for column in range(1, table_count * 2 + 2, 2)
        )
        zone = region_module._SparseZone.create(
            region_module._RunView.complete(runs),
            spans,
        )
        assert zone is not None
        tables = tuple(Rect(1, 12, column, column) for column in range(2, table_count * 2 + 1, 2))
        run_start = copied_runs
        span_start = constructed_spans

        children = region_module._partition_full_height_tables(zone, tables)

        assert len(children) == table_count + 1
        assert all(
            child.runs is not None and len(child.runs.materialize()) == 1 and len(child.spans) == 1
            for child in children
        )
        return copied_runs - run_start, constructed_spans - span_start

    small = batch_work(400)
    large = batch_work(800)

    assert small == (401, 401)
    assert large == (801, 801)


def test_blocked_near_components_filter_member_pair_candidates(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_build = region_module._build_sparse_member_block
    original_connect = region_module._sparse_primitives_connect
    original_nearby = region_module._nearby_sparse_components
    original_raw = region_module._sparse_root_member_candidates
    block_builds = 0
    indexed_members = 0
    primitive_candidates = 0
    component_candidates = 0
    raw_candidates = 0

    def counted_build(
        member_ids: tuple[int, ...],
        rectangles: Sequence[Rect],
    ) -> region_module._SparseMemberBlock:
        nonlocal block_builds, indexed_members
        block_builds += 1
        indexed_members += len(member_ids)
        return original_build(member_ids, rectangles)

    def counted_connect(
        left: Rect,
        right: Rect,
        *,
        gap_tol: int,
        table_barriers: region_module._RectangleIndex,
        corridor_cache: dict[tuple[str, int, int, int, int], bool],
    ) -> bool:
        nonlocal primitive_candidates
        primitive_candidates += 1
        return original_connect(
            left,
            right,
            gap_tol=gap_tol,
            table_barriers=table_barriers,
            corridor_cache=corridor_cache,
        )

    def counted_nearby(
        component_index: region_module._RectangleIndex,
        bounds: Rect,
        *,
        gap_tol: int,
        exclude: int,
    ) -> tuple[int, ...]:
        nonlocal component_candidates
        candidates = original_nearby(
            component_index,
            bounds,
            gap_tol=gap_tol,
            exclude=exclude,
        )
        component_candidates += len(candidates)
        return candidates

    def counted_raw(
        blocks: Sequence[region_module._SparseMemberBlock],
        query: Rect,
    ) -> tuple[int, ...]:
        nonlocal raw_candidates
        candidates = original_raw(blocks, query)
        raw_candidates += len(candidates)
        return candidates

    monkeypatch.setattr(region_module, "_build_sparse_member_block", counted_build)
    monkeypatch.setattr(region_module, "_sparse_primitives_connect", counted_connect)
    monkeypatch.setattr(region_module, "_nearby_sparse_components", counted_nearby)
    monkeypatch.setattr(region_module, "_sparse_root_member_candidates", counted_raw)

    def blocked_work(size: int) -> int:
        nonlocal block_builds
        nonlocal component_candidates
        nonlocal indexed_members
        nonlocal primitive_candidates
        nonlocal raw_candidates
        height = 2 * size - 1
        runs = tuple(
            run for row in range(1, height + 1, 2) for run in (_Run(row, 1, 1), _Run(row, 3, 3))
        )
        tables = (Rect(1, height, 2, 2),)
        build_start = block_builds
        primitive_start = primitive_candidates
        component_start = component_candidates
        indexed_start = indexed_members
        raw_start = raw_candidates

        components = region_module._closed_sparse_components(
            runs,
            (),
            _RectangleIndex(tables),
            gap_tol=1,
        )

        assert [
            (component.bounds, len(component.runs), len(component.spans))
            for component in components
        ] == [
            (Rect(1, height, 1, 1), size, 0),
            (Rect(1, height, 3, 3), size, 0),
        ]
        primitive_count = len(runs)
        built = block_builds - build_start
        indexed = indexed_members - indexed_start
        assert built <= 2 * primitive_count
        assert indexed <= primitive_count * primitive_count.bit_length()
        return (
            built
            + indexed
            + primitive_candidates
            - primitive_start
            + component_candidates
            - component_start
            + raw_candidates
            - raw_start
        )

    small = blocked_work(320)
    large = blocked_work(640)

    assert small <= 40 * 320
    assert large <= 40 * 640
    assert large <= small * 2.2


def test_root_local_member_blocks_avoid_global_same_root_noise(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_build = region_module._build_sparse_member_block
    original_connect = region_module._sparse_primitives_connect
    original_intersecting = region_module._intersecting_sparse_components
    original_nearby = region_module._nearby_sparse_components
    original_raw = region_module._sparse_root_member_candidates
    original_rectangles_intersect = region_module._rectangles_intersect
    block_builds = 0
    indexed_members = 0
    primitive_candidates = 0
    root_candidates = 0
    raw_candidates = 0
    spatial_checks = 0

    def counted_build(
        member_ids: tuple[int, ...],
        rectangles: Sequence[Rect],
    ) -> region_module._SparseMemberBlock:
        nonlocal block_builds, indexed_members
        block_builds += 1
        indexed_members += len(member_ids)
        return original_build(member_ids, rectangles)

    def counted_connect(
        left: Rect,
        right: Rect,
        *,
        gap_tol: int,
        table_barriers: region_module._RectangleIndex,
        corridor_cache: dict[tuple[str, int, int, int, int], bool],
    ) -> bool:
        nonlocal primitive_candidates
        primitive_candidates += 1
        return original_connect(
            left,
            right,
            gap_tol=gap_tol,
            table_barriers=table_barriers,
            corridor_cache=corridor_cache,
        )

    def counted_intersecting(
        component_index: region_module._RectangleIndex,
        bounds: Rect,
        *,
        exclude: int,
    ) -> tuple[int, ...]:
        nonlocal root_candidates
        candidates = original_intersecting(
            component_index,
            bounds,
            exclude=exclude,
        )
        root_candidates += len(candidates)
        return candidates

    def counted_nearby(
        component_index: region_module._RectangleIndex,
        bounds: Rect,
        *,
        gap_tol: int,
        exclude: int,
    ) -> tuple[int, ...]:
        nonlocal root_candidates
        candidates = original_nearby(
            component_index,
            bounds,
            gap_tol=gap_tol,
            exclude=exclude,
        )
        root_candidates += len(candidates)
        return candidates

    def counted_raw(
        blocks: Sequence[region_module._SparseMemberBlock],
        query: Rect,
    ) -> tuple[int, ...]:
        nonlocal raw_candidates
        candidates = original_raw(blocks, query)
        raw_candidates += len(candidates)
        return candidates

    def counted_rectangles_intersect(left: Rect, right: Rect) -> bool:
        nonlocal spatial_checks
        spatial_checks += 1
        return original_rectangles_intersect(left, right)

    monkeypatch.setattr(region_module, "_build_sparse_member_block", counted_build)
    monkeypatch.setattr(region_module, "_sparse_primitives_connect", counted_connect)
    monkeypatch.setattr(
        region_module,
        "_intersecting_sparse_components",
        counted_intersecting,
    )
    monkeypatch.setattr(region_module, "_nearby_sparse_components", counted_nearby)
    monkeypatch.setattr(region_module, "_sparse_root_member_candidates", counted_raw)
    monkeypatch.setattr(
        region_module,
        "_rectangles_intersect",
        counted_rectangles_intersect,
    )

    def strong_root_work(size: int) -> tuple[int, int]:
        nonlocal block_builds
        nonlocal indexed_members
        nonlocal primitive_candidates
        nonlocal raw_candidates
        nonlocal root_candidates
        nonlocal spatial_checks
        height = 2 * size - 1
        table_col = 2 * size
        runs = (
            *(_Run(1 + 2 * index, 1, table_col - 1) for index in range(size)),
            *(_Run(1 + 2 * index, table_col + 1, table_col + 1) for index in range(size)),
        )
        spans = tuple(
            region_module._Span(
                1,
                height,
                1 + 2 * index,
                1 + 2 * index,
            )
            for index in range(size)
        )
        build_start = block_builds
        indexed_start = indexed_members
        primitive_start = primitive_candidates
        raw_start = raw_candidates
        root_start = root_candidates
        spatial_start = spatial_checks

        components = region_module._closed_sparse_components(
            runs,
            spans,
            _RectangleIndex((Rect(1, height, table_col, table_col),)),
            gap_tol=1,
        )

        assert [
            (component.bounds, len(component.runs), len(component.spans))
            for component in components
        ] == [
            (Rect(1, height, 1, table_col - 1), size, size),
            (Rect(1, height, table_col + 1, table_col + 1), size, 0),
        ]
        primitive_count = len(runs) + len(spans)
        built = block_builds - build_start
        indexed = indexed_members - indexed_start
        assert built <= 2 * primitive_count
        assert indexed <= primitive_count * primitive_count.bit_length()
        checked = spatial_checks - spatial_start
        work = (
            built
            + indexed
            + primitive_candidates
            - primitive_start
            + raw_candidates
            - raw_start
            + root_candidates
            - root_start
            + checked
        )
        return work, checked

    small = strong_root_work(160)
    large = strong_root_work(320)

    assert small[0] <= 112 * 160
    assert large[0] <= 112 * 320
    assert large[0] <= small[0] * 2.2
    assert large[1] <= small[1] * 2.25


def test_connected_wraparound_partition_is_output_sensitive(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A connected wraparound may emit M-by-T pieces, but work tracks that output."""
    original_split = region_module._split_sparse_zone_around_table
    original_batch = region_module._partition_full_height_tables
    original_build = region_module._build_sparse_member_block
    original_intersecting = region_module._intersecting_sparse_components
    original_nearby = region_module._nearby_sparse_components
    original_raw = region_module._sparse_root_member_candidates
    original_connect = region_module._sparse_primitives_connect
    original_ordered = region_module._RectangleIndex.ordered_intersections
    original_intersections = region_module._RectangleIndex.intersections
    original_rectangles_intersect = region_module._rectangles_intersect
    active_barriers: region_module._RectangleIndex | None = None
    splitter_visits = 0
    child_fragments = 0
    block_builds = 0
    component_candidates = 0
    indexed_members = 0
    proximity_candidates = 0
    raw_member_candidates = 0
    spatial_checks = 0
    table_candidates = 0

    def fragment_count(zones: Sequence[region_module._SparseZone]) -> int:
        return sum(
            (0 if zone.runs is None else zone.runs.stop - zone.runs.start) + len(zone.spans)
            for zone in zones
        )

    def counted_split(
        zone: region_module._SparseZone,
        table: Rect,
    ) -> tuple[region_module._SparseZone, ...]:
        nonlocal splitter_visits, child_fragments
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
            splitter_visits += middle_stop - middle_start
        splitter_visits += len(zone.spans)
        children = original_split(zone, table)
        child_fragments += fragment_count(children)
        return children

    def counted_batch(
        zone: region_module._SparseZone,
        tables: Sequence[Rect],
    ) -> tuple[region_module._SparseZone, ...]:
        nonlocal splitter_visits, child_fragments
        splitter_visits += (0 if zone.runs is None else zone.runs.stop - zone.runs.start) + len(
            zone.spans
        )
        children = original_batch(zone, tables)
        child_fragments += fragment_count(children)
        return children

    def counted_connect(
        left: Rect,
        right: Rect,
        *,
        gap_tol: int,
        table_barriers: region_module._RectangleIndex,
        corridor_cache: dict[tuple[str, int, int, int, int], bool],
    ) -> bool:
        nonlocal proximity_candidates
        proximity_candidates += 1
        return original_connect(
            left,
            right,
            gap_tol=gap_tol,
            table_barriers=table_barriers,
            corridor_cache=corridor_cache,
        )

    def counted_build(
        member_ids: tuple[int, ...],
        rectangles: Sequence[Rect],
    ) -> region_module._SparseMemberBlock:
        nonlocal block_builds, indexed_members
        block_builds += 1
        indexed_members += len(member_ids)
        return original_build(member_ids, rectangles)

    def counted_intersecting(
        component_index: region_module._RectangleIndex,
        bounds: Rect,
        *,
        exclude: int,
    ) -> tuple[int, ...]:
        nonlocal component_candidates
        candidates = original_intersecting(
            component_index,
            bounds,
            exclude=exclude,
        )
        component_candidates += len(candidates)
        return candidates

    def counted_nearby(
        component_index: region_module._RectangleIndex,
        bounds: Rect,
        *,
        gap_tol: int,
        exclude: int,
    ) -> tuple[int, ...]:
        nonlocal component_candidates
        candidates = original_nearby(
            component_index,
            bounds,
            gap_tol=gap_tol,
            exclude=exclude,
        )
        component_candidates += len(candidates)
        return candidates

    def counted_raw(
        blocks: Sequence[region_module._SparseMemberBlock],
        query: Rect,
    ) -> tuple[int, ...]:
        nonlocal raw_member_candidates
        candidates = original_raw(blocks, query)
        raw_member_candidates += len(candidates)
        return candidates

    def counted_rectangles_intersect(left: Rect, right: Rect) -> bool:
        nonlocal spatial_checks
        spatial_checks += 1
        return original_rectangles_intersect(left, right)

    def counted_ordered(
        index: region_module._RectangleIndex,
        rect: Rect,
    ) -> Iterable[int]:
        candidates = original_ordered(index, rect)
        if index is not active_barriers:
            return candidates

        def counted_candidates() -> Iterable[int]:
            nonlocal table_candidates
            for candidate in candidates:
                table_candidates += 1
                yield candidate

        return counted_candidates()

    def counted_intersections(
        index: region_module._RectangleIndex,
        rect: Rect,
        *,
        exclude: int | None = None,
    ) -> tuple[int, ...]:
        nonlocal table_candidates
        candidates = original_intersections(index, rect, exclude=exclude)
        if index is active_barriers:
            table_candidates += len(candidates)
        return candidates

    monkeypatch.setattr(
        region_module,
        "_split_sparse_zone_around_table",
        counted_split,
    )
    monkeypatch.setattr(region_module, "_partition_full_height_tables", counted_batch)
    monkeypatch.setattr(region_module, "_build_sparse_member_block", counted_build)
    monkeypatch.setattr(
        region_module,
        "_intersecting_sparse_components",
        counted_intersecting,
    )
    monkeypatch.setattr(region_module, "_nearby_sparse_components", counted_nearby)
    monkeypatch.setattr(region_module, "_sparse_root_member_candidates", counted_raw)
    monkeypatch.setattr(region_module, "_sparse_primitives_connect", counted_connect)
    monkeypatch.setattr(
        region_module,
        "_rectangles_intersect",
        counted_rectangles_intersect,
    )
    monkeypatch.setattr(
        region_module._RectangleIndex,
        "ordered_intersections",
        counted_ordered,
    )
    monkeypatch.setattr(
        region_module._RectangleIndex,
        "intersections",
        counted_intersections,
    )

    def wraparound_work(size: int) -> tuple[int, int]:
        nonlocal active_barriers
        nonlocal block_builds
        nonlocal child_fragments
        nonlocal component_candidates
        nonlocal indexed_members
        nonlocal proximity_candidates
        nonlocal raw_member_candidates
        nonlocal spatial_checks
        nonlocal splitter_visits
        nonlocal table_candidates
        height = 2 * size + 1
        table_col = 3 * size
        runs = tuple(_Run(2 * index - 1, 1, table_col + 1) for index in range(1, size + 1))
        spans = (
            *(
                region_module._Span(
                    1,
                    height,
                    1 + 3 * index,
                    1 + 3 * index,
                )
                for index in range(size)
            ),
            region_module._Span(1, height, table_col + 1, table_col + 1),
        )
        tables = tuple(
            Rect(2 * index, 2 * index, table_col, table_col) for index in range(1, size + 1)
        )
        active_barriers = _RectangleIndex(tables)
        build_start = block_builds
        splitter_start = splitter_visits
        fragment_start = child_fragments
        component_start = component_candidates
        indexed_start = indexed_members
        proximity_start = proximity_candidates
        raw_start = raw_member_candidates
        spatial_start = spatial_checks
        table_start = table_candidates

        result = region_module._component_first_sparse_rectangles(
            runs,
            spans,
            tables,
            active_barriers,
            gap_tol=1,
        )

        primitive_count = 2 * size + 1
        expected_fragments = 7 * size * (size + 1) // 2
        expected_regions = size * size + 3 * size + 1
        total_fragments = (7 * size * size + 11 * size + 2) // 2
        assert len(runs) + len(spans) == primitive_count
        assert child_fragments - fragment_start == expected_fragments
        assert primitive_count + expected_fragments == total_fragments
        assert len(result) == expected_regions
        work = (
            splitter_visits
            - splitter_start
            + child_fragments
            - fragment_start
            + block_builds
            - build_start
            + component_candidates
            - component_start
            + indexed_members
            - indexed_start
            + proximity_candidates
            - proximity_start
            + raw_member_candidates
            - raw_start
            + spatial_checks
            - spatial_start
            + table_candidates
            - table_start
        )
        assert work <= 64 * (primitive_count + expected_fragments + expected_regions)
        return len(result), work

    small = wraparound_work(40)
    large = wraparound_work(80)

    assert large[1] <= small[1] * 4.5


def test_lazy_merge_spans_preserve_row_and_column_gap_connectivity() -> None:
    vertical_merge = Rect(1, 10, 1, 1)
    cells = [
        _value_cell(1, 1, "merge anchor"),
        _value_cell(5, 3, "column gap"),
        _value_cell(5, 4, "adjacent"),
        _value_cell(12, 2, "row gap"),
    ]

    gap_one = _analyze(
        cells,
        merges=(vertical_merge,),
        options=RegionOptions(gap_tol=1),
    )
    gap_zero = _analyze(
        cells,
        merges=(vertical_merge,),
        options=RegionOptions(gap_tol=0),
    )

    assert [region.rect for region in gap_one.regions] == [Rect(1, 12, 1, 4)]
    assert [region.rect for region in gap_zero.regions] == [
        Rect(1, 10, 1, 1),
        Rect(5, 5, 3, 4),
        Rect(12, 12, 2, 2),
    ]


def test_lazy_merge_components_keep_bounding_box_closure() -> None:
    merges = (
        Rect(1, 5, 1, 1),
        Rect(1, 1, 2, 5),
    )

    analysis = _analyze(
        [
            _value_cell(1, 1, "vertical anchor"),
            _value_cell(1, 2, "horizontal anchor"),
            _value_cell(5, 5, "closure"),
        ],
        merges=merges,
        options=RegionOptions(gap_tol=0),
    )

    assert [region.rect for region in analysis.regions] == [Rect(1, 5, 1, 5)]


def test_merge_row_coalescing_keeps_tolerated_gaps_and_table_barriers() -> None:
    merge = Rect(1, 2, 5, 5)
    cells = [
        _value_cell(1, 1, "left"),
        _value_cell(1, 5, "merge anchor"),
    ]
    without_table = _analyze(
        cells,
        merges=(merge,),
        options=RegionOptions(gap_tol=3),
    )
    barrier = TableInfo("Barrier", "Barrier", "C1:C2", 1, 0, ("Inside",))
    with_table = _analyze(
        [*cells, _value_cell(1, 3, "table")],
        tables=(barrier,),
        merges=(merge,),
        options=RegionOptions(gap_tol=3),
    )

    assert [region.rect for region in without_table.regions] == [Rect(1, 2, 1, 5)]
    assert [(region.kind, region.rect) for region in with_table.regions] == [
        ("region", Rect(1, 1, 1, 1)),
        ("table", Rect(1, 2, 3, 3)),
        ("region", Rect(1, 2, 5, 5)),
    ]


def test_many_same_row_merged_headers_reuse_sparse_lookup() -> None:
    region_count = 512
    merges = _CountingMerges(
        tuple(
            Rect(1, 1, 1 + region_index * 3, 2 + region_index * 3)
            for region_index in range(region_count)
        )
    )
    cells: list[RegionCell] = []
    for region_index in range(region_count):
        col = 1 + region_index * 3
        cells.extend(
            (
                _value_cell(1, col, f"Group {region_index}", style_idx=1),
                _value_cell(2, col, "Left", style_idx=1),
                _value_cell(2, col + 1, "Right", style_idx=1),
                _value_cell(3, col, region_index),
                _value_cell(3, col + 1, region_index + 1),
            )
        )
    ordered = tuple(sorted(cells, key=lambda cell: (cell.row, cell.col)))

    analysis = analyze_sheet_regions(
        _summary(
            ordered,
            merges=cast(tuple[Rect, ...], merges),
        ),
        _styles(),
        lambda: iter(ordered),
        RegionOptions(gap_tol=0),
    )

    assert len(analysis.regions) == region_count
    assert [column.header for column in analysis.regions[0].columns] == [
        "Group 0 / Left",
        "Group 0 / Right",
    ]
    assert merges.visits <= region_count * 4


def test_cascading_and_dense_rectangle_closure_has_bounded_candidate_work(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_intersects = region_module._rectangles_intersect
    checks = 0

    def counted_intersects(left: Rect, right: Rect) -> bool:
        nonlocal checks
        checks += 1
        return original_intersects(left, right)

    monkeypatch.setattr(region_module, "_rectangles_intersect", counted_intersects)

    staircase_count = 512
    staircase = [
        Rect(1, 1, 1, 2),
        Rect(1, 2, 1, 1),
        *(
            Rect(2 if index % 2 == 0 else 1, 2 if index % 2 == 0 else 1, index, index + 1)
            for index in range(2, staircase_count)
        ),
    ]
    start = checks
    staircase_result = _merge_intersecting_rectangles(staircase)
    staircase_work = checks - start

    dense_count = 512
    start = checks
    dense_result = _merge_intersecting_rectangles(
        tuple(Rect(10, 20, 10, 20) for _index in range(dense_count))
    )
    dense_work = checks - start
    empty_center_result = _merge_intersecting_rectangles(
        (
            Rect(1, 1, 1, 10),
            Rect(1, 10, 1, 1),
            Rect(15, 15, 5, 15),
            Rect(5, 15, 15, 15),
        )
    )

    assert staircase_result == (Rect(1, 2, 1, staircase_count),)
    assert dense_result == (Rect(10, 20, 10, 20),)
    assert empty_center_result == (Rect(1, 15, 1, 15),)
    assert staircase_work < staircase_count * 8
    assert dense_work < dense_count * 3


def test_numeric_grid_has_no_inferred_header_and_invalid_styles_are_safe() -> None:
    cells = [
        _value_cell(row, col, row * col, style_idx=999)
        for row in range(1, 4)
        for col in range(1, 3)
    ]

    region = _analyze(cells).regions[0]

    assert region.header_rows == 0
    assert [column.header for column in region.columns] == ["Column A", "Column B"]
    assert [column.nonnull for column in region.columns] == [3, 3]


def test_dtype_matrix_and_duplicate_header_ids() -> None:
    headers = ("Value", "Value!", "When", "Text", "Flag", "Mixed", "No Cache", "Errors")
    cells = [_value_cell(1, index, header) for index, header in enumerate(headers, start=1)]
    cells.extend(
        (
            _value_cell(2, 1, 1),
            _value_cell(3, 1, 2),
            _value_cell(2, 2, 1.5),
            _value_cell(3, 2, 2),
            RegionCell(2, 3, date(2026, 1, 1), "date"),
            RegionCell(3, 3, date(2026, 1, 2), "date"),
            _value_cell(2, 4, "A"),
            _value_cell(3, 4, "B"),
            _value_cell(2, 5, True),
            _value_cell(3, 5, False),
            _value_cell(2, 6, 1),
            _value_cell(3, 6, "one"),
            RegionCell(2, 7, None, "blank", formula="=A2"),
            RegionCell(3, 7, None, "blank", formula="=A3"),
            RegionCell(2, 8, "#DIV/0!", "error"),
            RegionCell(3, 8, "#N/A", "error"),
        )
    )

    columns = _analyze(cells).regions[0].columns

    assert [column.norm_header for column in columns[:2]] == ["value", "value#2"]
    assert [column.dtype for column in columns] == [
        "int",
        "float",
        "date",
        "str",
        "bool",
        "mixed",
        "empty",
        "mixed",
    ]
    assert columns[6].nonnull == 0


def test_distinct_and_dtype_samples_are_bounded_deterministically() -> None:
    cells = [_value_cell(1, 1, "Header")]
    cells.extend(_value_cell(row, 1, row) for row in range(2, 8))

    column = (
        _analyze(
            cells,
            options=RegionOptions(dtype_sample_limit=1, distinct_cap=3),
        )
        .regions[0]
        .columns[0]
    )

    assert column.dtype == "int"
    assert column.nonnull == 6
    assert column.distinct_est == 3


def test_large_sheet_warning_and_lower_rate_sample_are_deterministic() -> None:
    cells = [
        _value_cell(1, 1, "Header"),
        _value_cell(2, 1, 1),
        _value_cell(3, 1, "two"),
        _value_cell(4, 1, 3),
    ]
    options = RegionOptions(large_sheet_threshold=3, large_dtype_sample_limit=1)

    analysis = _analyze(cells, options=options)

    assert analysis.warnings[0].code == "W_LARGE_SHEET"
    assert analysis.warnings[0].ref == "sheet:Data"
    assert analysis.warnings[0].related == {
        "cellCount": 4,
        "dtypeSampleStride": 2,
        "dtypeSampleLimit": 1,
    }
    assert analysis.regions[0].columns[0].dtype == "int"


def test_analysis_rejects_duplicate_or_out_of_order_cells() -> None:
    cells = (
        _value_cell(2, 1, 1),
        _value_cell(1, 1, 2),
    )

    with pytest.raises(ValueError, match="strictly ordered"):
        analyze_sheet_regions(_summary(cells), _styles(), lambda: iter(cells))
