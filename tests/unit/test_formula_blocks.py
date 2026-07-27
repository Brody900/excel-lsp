"""Known-answer tests for R1C1 geometry and formula blocks."""

from __future__ import annotations

import pytest

import excel_lsp.core.formulas.tokens as tokens_module
from excel_lsp.core.formulas.a1 import (
    AxisTerm,
    CellRef,
    ParsedA1Reference,
    ReferenceGeometry,
    extrude_reference,
    parse_a1_reference,
    parse_modern_a1_range,
    render_r1c1_reference,
    resolve_reference,
)
from excel_lsp.core.formulas.blocks import (
    FormulaCell,
    FormulaPattern,
    InconsistentFormula,
    build_formula_blocks,
    detect_inconsistent_formulas,
    normalize_formula_cells,
)
from excel_lsp.core.formulas.r1c1 import to_r1c1
from excel_lsp.core.models import Rect


@pytest.mark.parametrize(
    ("formula", "expected"),
    (
        ("=A1+$B2+C$3+$D$4", "=R[-1]C[-1]+RC2+R3C[+1]+R4C4"),
        ("=A1:B3", "=R[-1]C[-1]:R[+1]C"),
        ("=A:A", "=C[-1]:C[-1]"),
        ("=$A:$C", "=C1:C3"),
        ("=1:3", "=R[-1]:R[+1]"),
        (
            "='My Sheet'!C3:D4",
            "='My Sheet'!R[+1]C[+1]:R[+2]C[+2]",
        ),
        ("='Jan 24:Mar 24'!B2", "='Jan 24:Mar 24'!RC"),
        (
            "=SUM(INDEX(B:B,1):A5)",
            "=SUM(INDEX(C:C,1):R[+3]C[-1])",
        ),
        (
            "=SUM(A1:INDEX(B:B,5))",
            "=SUM(R[-1]C[-1]:INDEX(C:C,5))",
        ),
    ),
)
def test_to_r1c1_known_answers(formula: str, expected: str) -> None:
    assert to_r1c1(formula, CellRef(2, 2)) == expected


def test_degenerate_range_spelling_keeps_exact_block_identity() -> None:
    range_formula = "=A1:A1"
    cell_formula = "=A2"

    assert to_r1c1(range_formula, CellRef(1, 2)) == "=RC[-1]:RC[-1]"
    assert to_r1c1(cell_formula, CellRef(2, 2)) == "=RC[-1]"

    patterns = normalize_formula_cells(
        (
            FormulaCell(1, 2, range_formula),
            FormulaCell(2, 2, cell_formula),
        )
    )
    assert [block.rect for block in build_formula_blocks(patterns)] == [
        Rect(1, 1, 2, 2),
        Rect(2, 2, 2, 2),
    ]

    assert to_r1c1("=A1:Rate", CellRef(1, 2)) == "=RC[-1]:Rate"
    assert to_r1c1("=A2:A2:Rate", CellRef(2, 2)) == "=RC[-1]:RC[-1]:Rate"


def test_block_growth_requires_exact_directional_translation() -> None:
    def blocks(first: str, second: str) -> tuple[Rect, ...]:
        patterns = normalize_formula_cells(
            (
                FormulaCell(1, 2, first),
                FormulaCell(2, 2, second),
            )
        )
        return tuple(block.rect for block in build_formula_blocks(patterns))

    def horizontal_blocks(first: str, second: str) -> tuple[Rect, ...]:
        patterns = normalize_formula_cells(
            (
                FormulaCell(1, 2, first),
                FormulaCell(1, 3, second),
            )
        )
        return tuple(block.rect for block in build_formula_blocks(patterns))

    assert blocks("=a1", "=A2") == (Rect(1, 2, 2, 2),)
    assert blocks("=@A1", "=@A2") == (Rect(1, 2, 2, 2),)
    assert blocks("=$a$1", "=$a$1") == (Rect(1, 2, 2, 2),)
    assert blocks("=A1", "=a2") == (Rect(1, 1, 2, 2), Rect(2, 2, 2, 2))
    assert blocks("=$a$1", "=$A$1") == (Rect(1, 1, 2, 2), Rect(2, 2, 2, 2))
    assert blocks("=A1#", "=A1#") == (Rect(1, 2, 2, 2),)
    assert blocks("=A1#", "=A2#") == (Rect(1, 1, 2, 2), Rect(2, 2, 2, 2))
    assert blocks("=A1#+A:A", "=A1#+A:A") == (Rect(1, 2, 2, 2),)
    assert blocks("=A1#+a:a", "=A1#+a:a") == (
        Rect(1, 1, 2, 2),
        Rect(2, 2, 2, 2),
    )
    assert horizontal_blocks("=a1", "=B1") == (Rect(1, 1, 2, 3),)
    assert horizontal_blocks("=A1", "=b1") == (
        Rect(1, 1, 2, 2),
        Rect(1, 1, 3, 3),
    )


def test_to_r1c1_preserves_non_reference_tokens_and_modern_operands() -> None:
    formula = '=SUM(  A1, "A1", Table1[Col], A1#, Name#, @A1)'
    assert to_r1c1(formula, CellRef(2, 2)) == (
        '=SUM(  R[-1]C[-1], "A1", Table1[Col], A1#, Name#, @R[-1]C[-1])'
    )


@pytest.mark.parametrize(
    ("formula", "expected"),
    (
        ("=A1#:B5", "=A1#:R[+3]C"),
        ("=A1:B5#", "=R[-1]C[-1]:B5#"),
        ("=@A1:B5", "=@R[-1]C[-1]:R[+3]C"),
        ("=A1:@B5", "=R[-1]C[-1]:@R[+3]C"),
        (
            "='Jan 24:Mar 24'!A1#:B5#",
            "='Jan 24:Mar 24'!A1#:B5#",
        ),
    ),
)
def test_to_r1c1_handles_modern_operators_on_either_range_endpoint(
    formula: str,
    expected: str,
) -> None:
    assert to_r1c1(formula, CellRef(2, 2)) == expected


@pytest.mark.parametrize(
    ("value", "qualifier"),
    (
        ("A1#:B5", ""),
        ("A1:@B5", ""),
        ("Jan:Mar!A1#:@B5", "Jan:Mar!"),
        ("@'Jan 24:Mar 24'!A1:B5#", "'Jan 24:Mar 24'!"),
        (
            "@[https://example.test/private/book.xlsx]Data!A1#:B5",
            "[https://example.test/private/book.xlsx]Data!",
        ),
    ),
)
def test_modern_range_parser_ignores_qualifier_and_bracket_colons(
    value: str,
    qualifier: str,
) -> None:
    parsed = parse_modern_a1_range(value, CellRef(2, 2))

    assert parsed is not None
    assert parsed.parsed.qualifier == qualifier


def test_modern_range_parser_does_not_split_a_structured_header_colon() -> None:
    assert parse_modern_a1_range("Table[@Header:Value]", CellRef(2, 2)) is None


def test_modern_range_normalization_preserves_spill_block_identity() -> None:
    spill_patterns = normalize_formula_cells(
        (
            FormulaCell(2, 2, "=A1#:B5"),
            FormulaCell(3, 2, "=A2#:B6"),
        )
    )
    intersection_patterns = normalize_formula_cells(
        (
            FormulaCell(2, 2, "=A1:@B5"),
            FormulaCell(3, 2, "=A2:@B6"),
        )
    )

    assert len(build_formula_blocks(spill_patterns)) == 2
    assert [pattern.r1c1 for pattern in intersection_patterns] == [
        "=R[-1]C[-1]:@R[+3]C",
        "=R[-1]C[-1]:@R[+3]C",
    ]
    assert [block.rect for block in build_formula_blocks(intersection_patterns)] == [
        Rect(2, 3, 2, 2)
    ]


@pytest.mark.parametrize(
    "value",
    (
        "Name",
        "R1C1",
        "XFE1",
        "A1048577",
        "Table1[Col]",
        "A1#",
        "@A1",
        "Sheet1!A1:Sheet1!B2",
    ),
)
def test_a1_parser_rejects_non_a1_or_unsupported_operands(value: str) -> None:
    assert parse_a1_reference(value, CellRef(2, 2)) is None


def test_a1_parser_preserves_qualifiers_and_absolute_bits() -> None:
    parsed = parse_a1_reference("'[budget.xlsx]Q1'!$A2:B$4", CellRef(3, 3))
    assert parsed is not None
    assert parsed == ParsedA1Reference(
        qualifier="'[budget.xlsx]Q1'!",
        geometry=ReferenceGeometry(
            row_a=AxisTerm(True, -1),
            row_b=AxisTerm(False, 4),
            col_a=AxisTerm(False, 1),
            col_b=AxisTerm(True, -1),
        ),
    )
    assert render_r1c1_reference(parsed) == "'[budget.xlsx]Q1'!R[-1]C1:R4C[-1]"
    assert resolve_reference(parsed.geometry, CellRef(3, 3)) == Rect(2, 4, 1, 2)


@pytest.mark.parametrize(
    ("token", "qualifier"),
    (
        ("A1", ""),
        ("Sheet2!A1", "Sheet2!"),
        ("'My Sheet'!A1", "'My Sheet'!"),
        ("'O''Brien'!A1", "'O''Brien'!"),
        ("Jan:Mar!A1", "Jan:Mar!"),
        ("'Jan 24:Mar 24'!A1", "'Jan 24:Mar 24'!"),
        ("[1]Sheet1!A1", "[1]Sheet1!"),
        ("'[budget.xlsx]Q1'!A1", "'[budget.xlsx]Q1'!"),
    ),
)
def test_a1_parser_retains_plain_sheet_3d_and_external_qualifiers(
    token: str,
    qualifier: str,
) -> None:
    parsed = parse_a1_reference(token, CellRef(2, 2))
    assert parsed is not None
    assert parsed.qualifier == qualifier
    assert resolve_reference(parsed.geometry, CellRef(2, 2)) == Rect(1, 1, 1, 1)


def test_extrusion_is_exact_for_relative_absolute_and_clamped_refs() -> None:
    anchor = CellRef(2, 2)
    relative = parse_a1_reference("A2", anchor)
    absolute = parse_a1_reference("$D$1", anchor)
    assert relative is not None and absolute is not None
    source = Rect(2, 50_001, 2, 2)
    assert extrude_reference(relative.geometry, source) == Rect(2, 50_001, 1, 1)
    assert extrude_reference(absolute.geometry, source) == Rect(1, 1, 4, 4)

    clipped = ReferenceGeometry(
        AxisTerm(True, 0),
        AxisTerm(True, 10),
        AxisTerm(True, 0),
        AxisTerm(True, 10),
    )
    assert extrude_reference(clipped, Rect(1_048_570, 1_048_576, 16_380, 16_384)) == Rect(
        1_048_570,
        1_048_576,
        16_380,
        16_384,
    )

    lower_clipped = ReferenceGeometry(
        AxisTerm(True, -1),
        AxisTerm(True, 0),
        AxisTerm(True, -1),
        AxisTerm(True, 0),
    )
    assert extrude_reference(lower_clipped, Rect(1, 2, 1, 2)) == Rect(1, 2, 1, 2)


def test_block_builder_grows_down_then_merges_equal_column_runs() -> None:
    cells = tuple(
        FormulaCell(row, col, f"={_column_label(col)}{row}")
        for col in (3, 4)
        for row in range(2, 6)
    )
    patterns = normalize_formula_cells(cells)
    blocks = build_formula_blocks(patterns)
    assert len(blocks) == 1
    assert blocks[0].n == 0
    assert blocks[0].rect == Rect(2, 5, 3, 4)
    assert blocks[0].r1c1 == "=RC"


def test_block_builder_groups_range_endpoints_absorbed_into_function_tokens() -> None:
    patterns = normalize_formula_cells(
        (
            FormulaCell(1, 4, "=SUM(A1:INDEX(B:B,5))"),
            FormulaCell(2, 4, "=SUM(A2:INDEX(B:B,5))"),
        )
    )

    blocks = build_formula_blocks(patterns)

    assert len(blocks) == 1
    assert blocks[0].rect == Rect(1, 2, 4, 4)
    assert blocks[0].r1c1 == "=SUM(RC[-3]:INDEX(C[-2]:C[-2],5))"


def test_fill_down_normalization_uses_the_exact_tokenizer_fast_path(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    compatibility_calls = 0
    original = tokens_module._shield_modern_syntax

    def counted(
        formula: str,
        spill_marker: str,
        implicit_qualifier_marker: str,
        structured_markers: dict[str, str],
    ) -> str:
        nonlocal compatibility_calls
        compatibility_calls += 1
        return original(
            formula,
            spill_marker,
            implicit_qualifier_marker,
            structured_markers,
        )

    monkeypatch.setattr(tokens_module, "_shield_modern_syntax", counted)
    cells = tuple(FormulaCell(row, 2, f"=A{row}*$J$2") for row in range(2, 5_002))

    patterns = normalize_formula_cells(cells)
    blocks = build_formula_blocks(patterns)

    assert compatibility_calls == 0
    assert len(blocks) == 1
    assert blocks[0].rect == Rect(2, 5_001, 2, 2)


@pytest.mark.parametrize(
    "formula",
    (
        "=@A2",
        "=A2:INDEX(A:A,2)",
    ),
)
def test_normalization_keeps_unsupported_translation_cases_distinct(
    formula: str,
) -> None:
    cells = (
        FormulaCell(2, 2, formula),
        FormulaCell(3, 2, formula),
    )

    patterns = normalize_formula_cells(cells)

    assert [pattern.r1c1 for pattern in patterns] == [
        to_r1c1(cell.formula, CellRef(cell.row, cell.col)) for cell in cells
    ]
    assert patterns[0].r1c1 != patterns[1].r1c1
    assert tuple(block.rect for block in build_formula_blocks(patterns)) == (
        Rect(2, 2, 2, 2),
        Rect(3, 3, 2, 2),
    )


def test_normalization_does_not_cross_the_last_excel_row() -> None:
    cells = (
        FormulaCell(1_048_575, 2, "=B1048576"),
        FormulaCell(1_048_576, 2, "=B1048577"),
    )

    patterns = normalize_formula_cells(cells)

    assert [pattern.r1c1 for pattern in patterns] == [
        "=R[+1]C",
        "=B1048577",
    ]


def test_block_builder_does_not_overcover_ragged_or_gapped_shapes() -> None:
    patterns = (
        *(FormulaPattern(row, 1, f"=A{row}", "=RC") for row in range(1, 4)),
        *(FormulaPattern(row, 2, f"=B{row}", "=RC") for row in range(1, 3)),
        FormulaPattern(1, 4, "=D1", "=RC"),
    )
    blocks = build_formula_blocks(patterns)
    assert tuple(block.rect for block in blocks) == (
        Rect(1, 3, 1, 1),
        Rect(1, 2, 2, 2),
        Rect(1, 1, 4, 4),
    )
    assert sum(_area(block.rect) for block in blocks) == len(patterns)


def test_normalization_keeps_malformed_formula_in_an_opaque_block() -> None:
    patterns = normalize_formula_cells((FormulaCell(1, 1, "=SUM(A1"),))
    assert patterns == (FormulaPattern(1, 1, "=SUM(A1", "=SUM(A1", parsed=False),)
    blocks = build_formula_blocks(patterns)
    assert blocks[0].parsed is False
    assert blocks[0].opaque is True
    assert blocks[0].rect == Rect(1, 1, 1, 1)


def test_duplicate_formula_coordinates_are_rejected() -> None:
    duplicate = FormulaPattern(1, 1, "=A1", "=RC")
    with pytest.raises(ValueError, match="unique"):
        build_formula_blocks((duplicate, duplicate))


def test_non_column_major_formula_patterns_are_rejected() -> None:
    patterns = (
        FormulaPattern(1, 2, "=B1", "=RC"),
        FormulaPattern(1, 1, "=A1", "=RC"),
    )
    with pytest.raises(ValueError, match="column-major"):
        build_formula_blocks(patterns)


def test_f07_tamper_is_the_only_inconsistency_and_uses_larger_dominant_block() -> None:
    cells = tuple(
        FormulaCell(
            row,
            3,
            f"=A{row}+B{row}" if row == 12 else f"=A{row}*B{row}",
        )
        for row in range(2, 22)
    )
    patterns = normalize_formula_cells(cells)
    blocks = build_formula_blocks(patterns)
    assert tuple(block.rect for block in blocks) == (
        Rect(2, 11, 3, 3),
        Rect(12, 12, 3, 3),
        Rect(13, 21, 3, 3),
    )
    assert detect_inconsistent_formulas(patterns, (Rect(1, 21, 1, 3),), blocks) == (
        InconsistentFormula(12, 3, "=RC[-2]*RC[-1]", 0),
    )


@pytest.mark.parametrize(
    ("run_length", "minorities", "flagged"),
    (
        (4, 1, False),
        (5, 1, True),
        (5, 2, False),
        (20, 1, True),
        (100, 4, True),
        (100, 6, False),
    ),
)
def test_inconsistency_threshold_boundaries(
    run_length: int,
    minorities: int,
    flagged: bool,
) -> None:
    patterns = tuple(
        FormulaPattern(
            row,
            1,
            "=1+1" if row <= minorities else "=A1",
            "=1+1" if row <= minorities else "=R[-1]C",
        )
        for row in range(1, run_length + 1)
    )
    blocks = build_formula_blocks(patterns)
    findings = detect_inconsistent_formulas(
        patterns,
        (Rect(1, run_length, 1, 1),),
        blocks,
    )
    assert bool(findings) is flagged
    if flagged:
        assert {(item.row, item.col) for item in findings} == {
            (row, 1) for row in range(1, minorities + 1)
        }


def test_inconsistency_runs_stop_at_gaps_and_region_boundaries() -> None:
    patterns = tuple(FormulaPattern(row, 1, "=A1", "=R[-1]C") for row in (1, 2, 3, 5, 6, 7))
    blocks = build_formula_blocks(patterns)
    assert (
        detect_inconsistent_formulas(
            patterns,
            (Rect(1, 3, 1, 1), Rect(5, 7, 1, 1)),
            blocks,
        )
        == ()
    )


def test_horizontal_and_vertical_findings_are_deduplicated() -> None:
    patterns = tuple(
        FormulaPattern(
            row,
            col,
            "=1+1" if (row, col) == (3, 3) else f"={_column_label(col)}{row}",
            "=1+1" if (row, col) == (3, 3) else "=RC",
        )
        for col in range(1, 6)
        for row in range(1, 6)
    )
    blocks = build_formula_blocks(patterns)
    findings = detect_inconsistent_formulas(patterns, (Rect(1, 5, 1, 5),), blocks)
    assert findings == (InconsistentFormula(3, 3, "=RC", 1),)


def _column_label(column: int) -> str:
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _area(rect: Rect) -> int:
    return (rect.row_max - rect.row_min + 1) * (rect.col_max - rect.col_min + 1)


@pytest.mark.parametrize(
    ("formula", "expected"),
    (
        ("=Rate:B5", "=Rate:R[+3]C"),
        ("=A1:rAtE", "=R[-1]C[-1]:rAtE"),
        ("=Rate#:B5", "=Rate#:R[+3]C"),
        ("=A1:Rate#", "=R[-1]C[-1]:Rate#"),
        ("=Rate:A1#", "=Rate:A1#"),
        ("=A1#:Rate", "=A1#:Rate"),
        ("=Start:End", "=Start:End"),
        ("=Start:(End)", "=Start:(End)"),
        ("=(Start):End", "=(Start):End"),
        ("=Start:B:B", "=Start:C:C"),
        ("=Start:5:5", "=Start:R[+3]:R[+3]"),
        ("=B:B:Start", "=C:C:Start"),
        ("=A:A:End", "=C[-1]:C[-1]:End"),
        ("=End:A:A", "=End:C[-1]:C[-1]"),
        (
            "='Jan 24:Mar 24'!A1:Rate",
            "='Jan 24:Mar 24'!R[-1]C[-1]:Rate",
        ),
        (
            "=Rate:'[https://example.test/a:b/book.xlsx]My Sheet'!B5",
            "=Rate:'[https://example.test/a:b/book.xlsx]My Sheet'!R[+3]C",
        ),
    ),
)
def test_to_r1c1_normalizes_only_the_a1_side_of_mixed_name_ranges(
    formula: str,
    expected: str,
) -> None:
    assert to_r1c1(formula, CellRef(2, 2)) == expected


@pytest.mark.parametrize(
    ("formula", "expected"),
    (
        ("=@'My Sheet'!A1:B5", "=@'My Sheet'!R[-1]C[-1]:R[+3]C"),
        ("=@('My Sheet'!A1:B5)", "=@('My Sheet'!R[-1]C[-1]:R[+3]C)"),
        ("=A1:@(B5)", "=R[-1]C[-1]:@(R[+3]C)"),
        ("=@(A1):B5", "=@(R[-1]C[-1]):R[+3]C"),
    ),
)
def test_to_r1c1_handles_qualified_and_grouped_implicit_intersections(
    formula: str,
    expected: str,
) -> None:
    assert to_r1c1(formula, CellRef(2, 2)) == expected
