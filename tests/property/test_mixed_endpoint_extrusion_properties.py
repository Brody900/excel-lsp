"""Differential properties for mixed relative and fixed range endpoints."""

from __future__ import annotations

from dataclasses import dataclass

from hypothesis import given, settings
from hypothesis import strategies as st

from excel_lsp.core.formulas.a1 import resolve_reference
from excel_lsp.core.formulas.analysis import analyze_formula
from excel_lsp.core.formulas.blocks import FormulaCell
from excel_lsp.core.formulas.indexing import analyze_sheet_formulas
from excel_lsp.core.formulas.references import FormulaAnchor, ReferenceContext, TableBinding
from excel_lsp.core.models import Rect, SheetDescriptor, TableInfo
from excel_lsp.core.parse.coordinates import column_label

_MAX_ROW = 1_048_576
_MAX_COLUMN = 16_384


@dataclass(frozen=True, slots=True)
class _MixedEndpointCase:
    source_min: int
    source_max: int
    source_col_min: int
    source_col_max: int
    endpoint_offset: int
    endpoint_col_offset: int
    fixed_min: int
    fixed_max: int
    fixed_col_min: int
    fixed_col_max: int
    reverse: bool


@st.composite
def _mixed_endpoint_cases(draw: st.DrawFn) -> _MixedEndpointCase:
    source_min = draw(
        st.one_of(
            st.integers(1, 25),
            st.integers(_MAX_ROW - 25, _MAX_ROW),
        )
    )
    maximum_height = min(8, _MAX_ROW - source_min + 1)
    height = draw(st.integers(1, maximum_height))
    source_max = source_min + height - 1
    source_col_min = draw(
        st.one_of(
            st.integers(1, 25),
            st.integers(_MAX_COLUMN - 25, _MAX_COLUMN),
        )
    )
    maximum_width = min(5, _MAX_COLUMN - source_col_min + 1)
    width = draw(st.integers(1, maximum_width))
    source_col_max = source_col_min + width - 1
    endpoint_offset = draw(
        st.integers(
            max(-6, 1 - source_min),
            min(6, _MAX_ROW - source_max),
        )
    )
    endpoint_col_offset = draw(
        st.integers(
            max(-6, 1 - source_col_min),
            min(6, _MAX_COLUMN - source_col_max),
        )
    )
    neighborhood_min = max(1, source_min - 8)
    neighborhood_max = min(_MAX_ROW, source_max + 8)
    fixed_min = draw(st.integers(neighborhood_min, neighborhood_max))
    fixed_max = draw(st.integers(fixed_min, neighborhood_max))
    column_neighborhood_min = max(1, source_col_min - 8)
    column_neighborhood_max = min(_MAX_COLUMN, source_col_max + 8)
    fixed_col_min = draw(st.integers(column_neighborhood_min, column_neighborhood_max))
    fixed_col_max = draw(st.integers(fixed_col_min, column_neighborhood_max))
    return _MixedEndpointCase(
        source_min,
        source_max,
        source_col_min,
        source_col_max,
        endpoint_offset,
        endpoint_col_offset,
        fixed_min,
        fixed_max,
        fixed_col_min,
        fixed_col_max,
        draw(st.booleans()),
    )


def _sheet() -> SheetDescriptor:
    return SheetDescriptor(
        order=0,
        name="Data",
        sheet_id=1,
        rel_id="rId1",
        xml_part="xl/worksheets/sheet1.xml",
        kind="worksheet",
    )


@given(_mixed_endpoint_cases())
@settings(max_examples=150, derandomize=True, database=None, deadline=None)
def test_mixed_endpoint_block_edge_matches_brute_per_cell_union(
    case: _MixedEndpointCase,
) -> None:
    sheet = _sheet()
    table = TableInfo(
        "FixedBand",
        "FixedBand",
        (
            f"{column_label(case.fixed_col_min)}{case.fixed_min}:"
            f"{column_label(case.fixed_col_max)}{case.fixed_max}"
        ),
        0,
        0,
        tuple(f"Field{offset}" for offset in range(case.fixed_col_max - case.fixed_col_min + 1)),
    )
    context = ReferenceContext(
        (sheet,),
        tables=(TableBinding(0, "Data", table),),
    )

    def formula(row: int, col: int) -> str:
        moving = f"{column_label(col + case.endpoint_col_offset)}{row + case.endpoint_offset}"
        fixed = "FixedBand[#Data]"
        left, right = (fixed, moving) if case.reverse else (moving, fixed)
        return f"=SUM({left}:{right})"

    cells = tuple(
        FormulaCell(
            row,
            col,
            formula(row, col),
        )
        for col in range(case.source_col_min, case.source_col_max + 1)
        for row in range(case.source_min, case.source_max + 1)
    )

    result = analyze_sheet_formulas(
        sheet,
        cells,
        (
            Rect(
                case.source_min,
                case.source_max,
                case.source_col_min,
                case.source_col_max,
            ),
        ),
        context,
    )

    per_cell_rects: list[Rect] = []
    for cell in cells:
        anchor = FormulaAnchor(0, "Data", cell.row, cell.col)
        semantic = analyze_formula(cell.formula, anchor=anchor, context=context)
        assert semantic.issues == ()
        assert not semantic.opaque
        assert len(semantic.references) == 1
        reference = semantic.references[0]
        assert reference.via == "structured:FixedBand"
        assert reference.geometry is not None
        per_cell_rects.append(resolve_reference(reference.geometry, anchor.cell))

    expected = Rect(
        min(rect.row_min for rect in per_cell_rects),
        max(rect.row_max for rect in per_cell_rects),
        min(rect.col_min for rect in per_cell_rects),
        max(rect.col_max for rect in per_cell_rects),
    )
    assert len(result.blocks) == 1
    assert result.edges[0].via == "structured:FixedBand"
    assert result.edges[0].rect == expected
    assert result.edges[0].rect is not None
    assert 1 <= result.edges[0].rect.row_min <= result.edges[0].rect.row_max <= _MAX_ROW
    assert 1 <= result.edges[0].rect.col_min <= result.edges[0].rect.col_max <= _MAX_COLUMN
