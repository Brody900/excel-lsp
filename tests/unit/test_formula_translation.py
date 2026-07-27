"""Known-answer tests for modern-safe shared-formula translation."""

from __future__ import annotations

import pytest

from excel_lsp.core.formulas.a1 import CellRef
from excel_lsp.core.formulas.translation import translate_a1_formula


def test_translates_relative_and_absolute_axes_without_changing_other_text() -> None:
    translated = translate_a1_formula(
        "=A2+$B2+C$2+$D$4",
        origin=CellRef(2, 2),
        target=CellRef(5, 4),
    )

    assert translated == "=C5+$B5+E$2+$D$4"


def test_translation_matches_excel_case_rules_for_absolute_columns() -> None:
    translated = translate_a1_formula(
        "=a1+$a1+a$1+$a$1+a:a+$a:$a",
        origin=CellRef(1, 2),
        target=CellRef(2, 2),
    )

    assert translated == "=A2+$a2+A$1+$a$1+A:A+$a:$a"


def test_block_translation_can_preserve_coordinate_spill_anchors() -> None:
    formula = "=A1#+a:a+A1#:@B5"

    assert (
        translate_a1_formula(
            formula,
            origin=CellRef(1, 2),
            target=CellRef(2, 2),
        )
        == "=A2#+A:A+A2#:@B6"
    )
    assert (
        translate_a1_formula(
            formula,
            origin=CellRef(1, 2),
            target=CellRef(2, 2),
            preserve_coordinate_spills=True,
        )
        == "=A1#+A:A+A1#:@B6"
    )


def test_translates_modern_operands_and_absorbed_range_endpoints() -> None:
    translated = translate_a1_formula(
        "=@A2+A2#+Name#+Esc[A'[B]+A2:INDEX(A:A,2)",
        origin=CellRef(2, 2),
        target=CellRef(3, 2),
    )

    assert translated == "=@A3+A3#+Name#+Esc[A'[B]+A3:INDEX(A:A,2)"


def test_preserves_qualifiers_range_shape_and_structured_references() -> None:
    translated = translate_a1_formula(
        "='My Sheet'!A2:A2+'[book.xlsx]Data'!$B2+Table1[@Value]",
        origin=CellRef(2, 2),
        target=CellRef(5, 4),
    )

    assert translated == "='My Sheet'!C5:C5+'[book.xlsx]Data'!$B5+Table1[@Value]"


def test_translates_modern_operators_on_both_range_endpoints() -> None:
    translated = translate_a1_formula(
        (
            "=A1#:B5+A1:@B5+@A1:B5#"
            "+'Jan 24:Mar 24'!A1#:@B5"
            "+'[https://example.test/private/book.xlsx]Data'!A1:@B5#"
        ),
        origin=CellRef(2, 2),
        target=CellRef(3, 2),
    )

    assert translated == (
        "=A2#:B6+A2:@B6+@A2:B6#"
        "+'Jan 24:Mar 24'!A2#:@B6"
        "+'[https://example.test/private/book.xlsx]Data'!A2:@B6#"
    )


def test_rejects_a_translation_outside_excel_bounds() -> None:
    with pytest.raises(ValueError, match="bounds"):
        translate_a1_formula(
            "=A1",
            origin=CellRef(2, 1),
            target=CellRef(1, 1),
        )


@pytest.mark.parametrize(
    ("formula", "expected"),
    (
        ("=Rate:B5", "=Rate:B6"),
        ("=A1:rAtE", "=A2:rAtE"),
        ("=Rate#:B5", "=Rate#:B6"),
        ("=A1:Rate#", "=A2:Rate#"),
        ("=Rate:A1#", "=Rate:A2#"),
        ("=A1#:Rate", "=A2#:Rate"),
        ("=Start:End", "=Start:End"),
        ("=Start:(End)", "=Start:(End)"),
        ("=(Start):End", "=(Start):End"),
        ("=Start:B:B", "=Start:B:B"),
        ("=Start:5:5", "=Start:6:6"),
        ("=B:B:Start", "=B:B:Start"),
        ("='Jan 24:Mar 24'!A1:Rate", "='Jan 24:Mar 24'!A2:Rate"),
        (
            "=Rate:'[https://example.test/a:b/book.xlsx]My Sheet'!B5",
            "=Rate:'[https://example.test/a:b/book.xlsx]My Sheet'!B6",
        ),
    ),
)
def test_translates_only_the_a1_side_of_mixed_name_ranges(
    formula: str,
    expected: str,
) -> None:
    assert (
        translate_a1_formula(
            formula,
            origin=CellRef(2, 2),
            target=CellRef(3, 2),
        )
        == expected
    )


@pytest.mark.parametrize(
    ("formula", "expected"),
    (
        ("=@'My Sheet'!A1:B5", "=@'My Sheet'!A2:B6"),
        ("=@('My Sheet'!A1:B5)", "=@('My Sheet'!A2:B6)"),
        ("=A1:@(B5)", "=A2:@(B6)"),
        ("=@(A1):B5", "=@(A2):B6"),
    ),
)
def test_translates_qualified_and_grouped_implicit_intersections(
    formula: str,
    expected: str,
) -> None:
    assert (
        translate_a1_formula(
            formula,
            origin=CellRef(2, 2),
            target=CellRef(3, 2),
        )
        == expected
    )


def test_translates_repeated_whole_axis_before_column_like_defined_name() -> None:
    assert (
        translate_a1_formula(
            "=A:A:End",
            origin=CellRef(2, 2),
            target=CellRef(2, 3),
        )
        == "=B:B:End"
    )
    assert (
        translate_a1_formula(
            "=End:A:A",
            origin=CellRef(2, 2),
            target=CellRef(2, 3),
        )
        == "=End:B:B"
    )
