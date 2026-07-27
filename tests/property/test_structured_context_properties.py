"""Differential properties for source-dependent bare structured references."""

from __future__ import annotations

from dataclasses import dataclass
from itertools import combinations

from hypothesis import given, settings
from hypothesis import strategies as st

from excel_lsp.core.formulas.a1 import resolve_reference
from excel_lsp.core.formulas.blocks import FormulaCell
from excel_lsp.core.formulas.indexing import (
    SheetFormulaAnalysis,
    _partition_context_rectangles,
    analyze_sheet_formulas,
)
from excel_lsp.core.formulas.references import (
    FormulaAnchor,
    ReferenceContext,
    TableBinding,
    classify_ref,
)
from excel_lsp.core.models import Rect, SheetDescriptor, TableInfo
from excel_lsp.core.parse.coordinates import column_label

_CoveragePoint = tuple[int | None, int | None, int | None, str]


@dataclass(frozen=True, slots=True)
class _TableGrid:
    sheet: SheetDescriptor
    tables: tuple[TableBinding, ...]
    table_rects: tuple[Rect, ...]
    formula_rect: Rect

    def formula_cells(self, formula: str) -> tuple[FormulaCell, ...]:
        return tuple(
            FormulaCell(row, col, formula)
            for col in range(self.formula_rect.col_min, self.formula_rect.col_max + 1)
            for row in range(self.formula_rect.row_min, self.formula_rect.row_max + 1)
        )


def _sheet() -> SheetDescriptor:
    return SheetDescriptor(
        order=0,
        name="Structured",
        sheet_id=1,
        rel_id="rId1",
        xml_part="xl/worksheets/sheet1.xml",
        kind="worksheet",
    )


@st.composite
def _table_grids(draw: st.DrawFn, *, include_boundaries: bool) -> _TableGrid:
    """Tile one bounded formula rectangle with randomly shaped ListObjects."""
    row_band_count, col_band_count = draw(st.sampled_from(((1, 2), (1, 3), (2, 1), (2, 2), (3, 1))))
    start_row = draw(st.integers(1, 8))
    start_col = draw(st.integers(1, 8))
    minimum_height = 3 if include_boundaries else 1
    row_heights = draw(
        st.lists(
            st.integers(minimum_height, 6),
            min_size=row_band_count,
            max_size=row_band_count,
        )
    )
    col_widths = draw(
        st.lists(
            st.integers(1, 4),
            min_size=col_band_count,
            max_size=col_band_count,
        )
    )

    row_starts = [start_row]
    for height in row_heights[:-1]:
        row_starts.append(row_starts[-1] + height)
    col_starts = [start_col]
    for width in col_widths[:-1]:
        col_starts.append(col_starts[-1] + width)

    sheet = _sheet()
    tables: list[TableBinding] = []
    table_rects: list[Rect] = []
    table_n = 0
    for row_start, height in zip(row_starts, row_heights, strict=True):
        for col_start, width in zip(col_starts, col_widths, strict=True):
            rect = Rect(
                row_start,
                row_start + height - 1,
                col_start,
                col_start + width - 1,
            )
            if include_boundaries and table_n == 0:
                header_rows, totals_rows = 1, 1
            elif include_boundaries:
                header_rows, totals_rows = draw(st.sampled_from(((0, 0), (1, 0), (0, 1), (1, 1))))
            else:
                header_rows, totals_rows = 0, 0

            value_offset = draw(st.integers(0, width - 1))
            columns = tuple(
                "Value" if offset == value_offset else f"Field{offset}" for offset in range(width)
            )
            name = f"Table{table_n}"
            table = TableInfo(
                name=name,
                display_name=name,
                ref=(
                    f"{column_label(rect.col_min)}{rect.row_min}:"
                    f"{column_label(rect.col_max)}{rect.row_max}"
                ),
                header_rows=header_rows,
                totals_rows=totals_rows,
                columns=columns,
            )
            tables.append(TableBinding(sheet.order, sheet.name, table))
            table_rects.append(rect)
            table_n += 1

    return _TableGrid(
        sheet=sheet,
        tables=tuple(tables),
        table_rects=tuple(table_rects),
        formula_rect=Rect(
            start_row,
            start_row + sum(row_heights) - 1,
            start_col,
            start_col + sum(col_widths) - 1,
        ),
    )


def _rect_points(
    sheet_order: int,
    rect: Rect,
    via: str,
) -> set[_CoveragePoint]:
    return {
        (sheet_order, row, col, via)
        for row in range(rect.row_min, rect.row_max + 1)
        for col in range(rect.col_min, rect.col_max + 1)
    }


def _analysis_coverage(result: SheetFormulaAnalysis) -> frozenset[_CoveragePoint]:
    coverage: set[_CoveragePoint] = set()
    for edge in result.edges:
        if edge.dst_sheet_order is None:
            assert edge.rect is None
            coverage.add((None, None, None, edge.via))
        else:
            assert edge.rect is not None
            coverage.update(_rect_points(edge.dst_sheet_order, edge.rect, edge.via))
    return frozenset(coverage)


def _brute_per_cell_coverage(
    layout: _TableGrid,
    token: str,
    context: ReferenceContext,
) -> frozenset[_CoveragePoint]:
    coverage: set[_CoveragePoint] = set()
    for col in range(layout.formula_rect.col_min, layout.formula_rect.col_max + 1):
        for row in range(layout.formula_rect.row_min, layout.formula_rect.row_max + 1):
            anchor = FormulaAnchor(layout.sheet.order, layout.sheet.name, row, col)
            classification = classify_ref(token, anchor=anchor, context=context)
            assert classification.issues == ()
            assert not classification.opaque
            for reference in classification.references:
                if reference.dst_sheet_order is None:
                    assert reference.geometry is None
                    coverage.add((None, None, None, reference.via))
                else:
                    assert reference.geometry is not None
                    rect = resolve_reference(reference.geometry, anchor.cell)
                    coverage.update(_rect_points(reference.dst_sheet_order, rect, reference.via))
    return frozenset(coverage)


def _expected_coverage_size(layout: _TableGrid, token: str) -> int:
    if token == "[@Value]":
        return sum(rect.row_max - rect.row_min + 1 for rect in layout.table_rects)
    return sum(
        (rect.row_max - rect.row_min + 1 - binding.table.header_rows - binding.table.totals_rows)
        * (rect.col_max - rect.col_min + 1)
        for binding, rect in zip(layout.tables, layout.table_rects, strict=True)
    )


def _assert_structured_differential(layout: _TableGrid, token: str) -> None:
    formula = f"={token}"
    context = ReferenceContext((layout.sheet,), tables=layout.tables)
    cells = layout.formula_cells(formula)

    result = analyze_sheet_formulas(layout.sheet, cells, (layout.formula_rect,), context)
    repeated = analyze_sheet_formulas(layout.sheet, cells, (layout.formula_rect,), context)

    assert result == repeated
    assert len(layout.tables) >= 2
    for left, right in combinations(layout.table_rects, 2):
        assert not left.intersects(right)
    assert [(block.n, block.rect, block.r1c1) for block in result.blocks] == [
        (0, layout.formula_rect, formula)
    ]
    assert not result.blocks[0].opaque
    assert result.diagnostics == ()
    assert {edge.source_block_n for edge in result.edges} == {0}

    brute = _brute_per_cell_coverage(layout, token, context)
    actual = _analysis_coverage(result)
    expected_vias = {
        f"structured:{binding.table.name}{'[Value]' if token == '[@Value]' else ''}"
        for binding in layout.tables
    }
    assert len(brute) == _expected_coverage_size(layout, token)
    assert {point[3] for point in brute} == expected_vias
    assert {point[3] for point in actual} == expected_vias
    assert actual == brute


@given(_table_grids(include_boundaries=False))
@settings(max_examples=40, derandomize=True, database=None, deadline=None)
def test_current_row_edges_match_brute_per_formula_cell_union(layout: _TableGrid) -> None:
    _assert_structured_differential(layout, "[@Value]")


@given(_table_grids(include_boundaries=True))
@settings(max_examples=40, derandomize=True, database=None, deadline=None)
def test_data_edges_match_brute_per_formula_cell_union_across_boundaries(
    layout: _TableGrid,
) -> None:
    first_table = layout.tables[0].table
    assert (first_table.header_rows, first_table.totals_rows) == (1, 1)
    _assert_structured_differential(layout, "[#Data]")


@st.composite
def _overlapping_contexts(
    draw: st.DrawFn,
) -> tuple[Rect, tuple[tuple[str, Rect], ...]]:
    height = draw(st.integers(1, 10))
    width = draw(st.integers(1, 10))
    bounds = Rect(1, height, 1, width)
    rectangles: list[tuple[str, Rect]] = []
    for _ in range(draw(st.integers(0, 15))):
        row_min = draw(st.integers(1, height))
        row_max = draw(st.integers(row_min, height))
        col_min = draw(st.integers(1, width))
        col_max = draw(st.integers(col_min, width))
        rectangles.append(
            (
                draw(st.sampled_from(("table:a", "table:b", "row:a", "row:b"))),
                Rect(row_min, row_max, col_min, col_max),
            )
        )
    return bounds, tuple(rectangles)


@given(_overlapping_contexts())
@settings(max_examples=200, derandomize=True, database=None, deadline=None)
def test_context_partition_is_exact_for_overlaps_and_duplicate_labels(
    case: tuple[Rect, tuple[tuple[str, Rect], ...]],
) -> None:
    bounds, contexts = case
    tiles = _partition_context_rectangles(bounds, contexts)
    owners: dict[tuple[int, int], int] = {}
    for tile_n, tile in enumerate(tiles):
        semantic_keys: set[tuple[str, ...]] = set()
        for row in range(tile.row_min, tile.row_max + 1):
            for col in range(tile.col_min, tile.col_max + 1):
                coordinate = (row, col)
                assert coordinate not in owners
                owners[coordinate] = tile_n
                semantic_keys.add(
                    tuple(
                        sorted(
                            {
                                label
                                for label, rect in contexts
                                if rect.row_min <= row <= rect.row_max
                                and rect.col_min <= col <= rect.col_max
                            }
                        )
                    )
                )
        assert len(semantic_keys) == 1

    assert set(owners) == {
        (row, col)
        for row in range(bounds.row_min, bounds.row_max + 1)
        for col in range(bounds.col_min, bounds.col_max + 1)
    }
