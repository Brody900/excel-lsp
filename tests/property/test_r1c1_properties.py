"""Property oracles for R1C1 normalization and exact formula rectangles."""

from __future__ import annotations

from itertools import combinations

from hypothesis import given, settings
from hypothesis import strategies as st
from openpyxl.formula.translate import Translator

from excel_lsp.core.formulas.a1 import (
    AxisTerm,
    CellRef,
    ReferenceGeometry,
    extrude_reference,
    resolve_reference,
)
from excel_lsp.core.formulas.blocks import (
    FormulaCell,
    build_formula_blocks,
    normalize_formula_cells,
)
from excel_lsp.core.formulas.r1c1 import to_r1c1
from excel_lsp.core.formulas.translation import translate_a1_formula
from excel_lsp.core.models import Rect
from excel_lsp.core.parse.coordinates import column_label, make_cell_ref


@st.composite
def _translation_cases(draw: st.DrawFn) -> tuple[str, CellRef, CellRef]:
    origin = CellRef(draw(st.integers(20, 100)), draw(st.integers(20, 100)))
    target = CellRef(draw(st.integers(20, 100)), draw(st.integers(20, 100)))
    qualifier = draw(st.sampled_from(("", "Sheet2!", "'My Sheet'!", "'Jan 24:Mar 24'!")))

    def reference() -> str:
        kind = draw(st.sampled_from(("cell", "range", "columns", "rows")))
        if kind in {"cell", "range"}:
            endpoint_count = 1 if kind == "cell" else 2
            endpoints: list[str] = []
            for _ in range(endpoint_count):
                column = draw(st.integers(100, 300))
                row = draw(st.integers(200, 500))
                col_abs = "$" if draw(st.booleans()) else ""
                row_abs = "$" if draw(st.booleans()) else ""
                endpoints.append(f"{col_abs}{column_label(column)}{row_abs}{row}")
            return qualifier + ":".join(endpoints)
        if kind == "columns":
            columns = draw(st.lists(st.integers(100, 300), min_size=2, max_size=2))
            markers = draw(st.lists(st.booleans(), min_size=2, max_size=2))
            rendered = (
                f"{'$' if absolute else ''}{column_label(column)}"
                for column, absolute in zip(columns, markers, strict=True)
            )
            return qualifier + ":".join(rendered)
        rows = draw(st.lists(st.integers(200, 500), min_size=2, max_size=2))
        markers = draw(st.lists(st.booleans(), min_size=2, max_size=2))
        rendered = (
            f"{'$' if absolute else ''}{row}" for row, absolute in zip(rows, markers, strict=True)
        )
        return qualifier + ":".join(rendered)

    first = reference()
    second = reference()
    shape = draw(st.sampled_from(("single", "sum", "binary")))
    if shape == "single":
        formula = f"={first}"
    elif shape == "sum":
        formula = f"=SUM({first},{second})"
    else:
        formula = f"={first}+{second}*2"
    return formula, origin, target


@given(_translation_cases())
@settings(max_examples=300, derandomize=True, database=None)
def test_i10_translation_invariance(case: tuple[str, CellRef, CellRef]) -> None:
    formula, origin, target = case
    shifted = Translator(
        formula,
        origin=make_cell_ref(origin.row, origin.col),
    ).translate_formula(make_cell_ref(target.row, target.col))
    assert to_r1c1(formula, origin) == to_r1c1(shifted, target)


@given(_translation_cases())
@settings(max_examples=300, derandomize=True, database=None)
def test_shared_formula_translation_matches_openpyxl_on_supported_a1_grammar(
    case: tuple[str, CellRef, CellRef],
) -> None:
    formula, origin, target = case
    expected = Translator(
        formula,
        origin=make_cell_ref(origin.row, origin.col),
    ).translate_formula(make_cell_ref(target.row, target.col))

    actual = translate_a1_formula(formula, origin=origin, target=target)

    assert actual == expected


def test_i10_translation_invariance_for_index_right_hand_range_endpoint() -> None:
    formula = "=SUM(INDEX(B:B,1):A5)"
    origin = CellRef(1, 4)
    target = CellRef(2, 4)
    shifted = Translator(formula, origin="D1").translate_formula("D2")

    assert shifted == "=SUM(INDEX(B:B,1):A6)"
    assert to_r1c1(formula, origin) == to_r1c1(shifted, target)


@st.composite
def _mixed_name_a1_cases(
    draw: st.DrawFn,
) -> tuple[str, str, CellRef, CellRef, bool]:
    origin = CellRef(draw(st.integers(20, 100)), draw(st.integers(20, 100)))
    target = CellRef(draw(st.integers(20, 100)), draw(st.integers(20, 100)))
    column = draw(st.integers(100, 300))
    row = draw(st.integers(200, 500))
    col_absolute = draw(st.booleans())
    row_absolute = draw(st.booleans())
    qualifier = draw(
        st.sampled_from(
            (
                "",
                "'My Sheet'!",
                "'Jan 24:Mar 24'!",
                "'[https://example.test/a:b/book.xlsx]Data'!",
            )
        )
    )
    a1_implicit = draw(st.booleans())
    a1_spill = draw(st.booleans())
    name_implicit = draw(st.booleans())
    name_spill = draw(st.booleans())
    name = draw(st.sampled_from(("Rate", "rAtE_2", "_Exact.Name")))
    name_on_left = draw(st.booleans())

    def coordinate(at: CellRef) -> str:
        translated_column = column if col_absolute else column + at.col - origin.col
        translated_row = row if row_absolute else row + at.row - origin.row
        return (
            f"{'$' if col_absolute else ''}{column_label(translated_column)}"
            f"{'$' if row_absolute else ''}{translated_row}"
        )

    def a1_endpoint(at: CellRef) -> str:
        return f"{'@' if a1_implicit else ''}{qualifier}{coordinate(at)}{'#' if a1_spill else ''}"

    name_endpoint = f"{'@' if name_implicit else ''}{name}{'#' if name_spill else ''}"
    original_parts = (
        (name_endpoint, a1_endpoint(origin))
        if name_on_left
        else (a1_endpoint(origin), name_endpoint)
    )
    translated_parts = (
        (name_endpoint, a1_endpoint(target))
        if name_on_left
        else (a1_endpoint(target), name_endpoint)
    )
    return (
        f"={original_parts[0]}:{original_parts[1]}",
        f"={translated_parts[0]}:{translated_parts[1]}",
        origin,
        target,
        a1_spill,
    )


@given(_mixed_name_a1_cases())
@settings(max_examples=250, derandomize=True, database=None)
def test_mixed_name_a1_translation_preserves_name_spelling_and_i10(
    case: tuple[str, str, CellRef, CellRef, bool],
) -> None:
    formula, expected, origin, target, a1_spill = case

    assert translate_a1_formula(formula, origin=origin, target=target) == expected
    if not a1_spill:
        assert to_r1c1(formula, origin) == to_r1c1(expected, target)


_GRID_ATOMS = st.sampled_from((None, "self", "left", "absolute", "constant"))


@st.composite
def _small_formula_grids(draw: st.DrawFn) -> tuple[FormulaCell, ...]:
    height = draw(st.integers(1, 7))
    width = draw(st.integers(1, 7))
    atoms = draw(st.lists(_GRID_ATOMS, min_size=height * width, max_size=height * width))
    cells: list[FormulaCell] = []
    for offset, atom in enumerate(atoms):
        if atom is None:
            continue
        row = offset // width + 1
        col = offset % width + 3
        if atom == "self":
            formula = f"={make_cell_ref(row, col)}"
        elif atom == "left":
            formula = f"={make_cell_ref(row, col - 1)}"
        elif atom == "absolute":
            formula = "=$A$1"
        else:
            formula = "=1+1"
        cells.append(FormulaCell(row, col, formula))
    return tuple(cells)


@given(_small_formula_grids())
@settings(max_examples=200, derandomize=True, database=None)
def test_i9_i11_formula_blocks_cover_exactly_once(cells: tuple[FormulaCell, ...]) -> None:
    column_major = sorted(cells, key=lambda cell: (cell.col, cell.row))
    patterns = normalize_formula_cells(column_major)
    blocks = build_formula_blocks(patterns)
    expected = {(cell.row, cell.col) for cell in cells}
    owners: dict[tuple[int, int], int] = {}
    for block in blocks:
        for row in range(block.rect.row_min, block.rect.row_max + 1):
            for col in range(block.rect.col_min, block.rect.col_max + 1):
                coordinate = (row, col)
                assert coordinate not in owners
                owners[coordinate] = block.n
                cell = next(item for item in cells if (item.row, item.col) == coordinate)
                assert to_r1c1(cell.formula, CellRef(row, col)) == block.r1c1
    assert set(owners) == expected
    assert build_formula_blocks(patterns) == blocks
    for left, right in combinations(blocks, 2):
        assert not left.rect.intersects(right.rect)


_AXIS_TERMS = st.one_of(
    st.builds(AxisTerm, st.just(True), st.integers(-10, 10)),
    st.builds(AxisTerm, st.just(False), st.integers(1, 100)),
)


@given(
    row=st.integers(20, 50),
    col=st.integers(20, 50),
    height=st.integers(1, 10),
    width=st.integers(1, 10),
    row_a=_AXIS_TERMS,
    row_b=_AXIS_TERMS,
    col_a=_AXIS_TERMS,
    col_b=_AXIS_TERMS,
)
@settings(max_examples=250, derandomize=True, database=None)
def test_extrusion_matches_brute_per_cell_union(
    row: int,
    col: int,
    height: int,
    width: int,
    row_a: AxisTerm,
    row_b: AxisTerm,
    col_a: AxisTerm,
    col_b: AxisTerm,
) -> None:
    geometry = ReferenceGeometry(row_a, row_b, col_a, col_b)
    source = Rect(row, row + height - 1, col, col + width - 1)
    resolved = tuple(
        resolve_reference(geometry, CellRef(source_row, source_col))
        for source_row in range(source.row_min, source.row_max + 1)
        for source_col in range(source.col_min, source.col_max + 1)
    )
    brute = Rect(
        min(rect.row_min for rect in resolved),
        max(rect.row_max for rect in resolved),
        min(rect.col_min for rect in resolved),
        max(rect.col_max for rect in resolved),
    )
    assert extrude_reference(geometry, source) == brute
