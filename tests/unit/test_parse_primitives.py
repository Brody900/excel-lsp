"""Known-answer tests for coordinate and SpreadsheetML style primitives."""

from __future__ import annotations

import lxml.etree as etree
import pytest

from excel_lsp.core.models import Rect
from excel_lsp.core.parse.coordinates import (
    column_label,
    column_number,
    contains,
    make_cell_ref,
    parse_cell_ref,
    parse_rect,
)
from excel_lsp.core.parse.styles import (
    BUILTIN_DATE_FORMAT_IDS,
    CellStyle,
    custom_format_is_date,
    parse_style_catalog,
)


@pytest.mark.parametrize(
    ("label", "number"),
    [("A", 1), ("z", 26), ("AA", 27), ("XFD", 16_384)],
)
def test_column_labels_round_trip_at_known_boundaries(label: str, number: int) -> None:
    assert column_number(label) == number
    assert column_label(number) == label.upper()


@pytest.mark.parametrize("label", ["", "A1", "A-A", "XFE"])
def test_column_number_rejects_invalid_or_out_of_bounds_labels(label: str) -> None:
    with pytest.raises(ValueError):
        column_number(label)


@pytest.mark.parametrize("column", [0, -1, 16_385])
def test_column_label_rejects_out_of_bounds_numbers(column: int) -> None:
    with pytest.raises(ValueError):
        column_label(column)


def test_cell_references_accept_absolutes_and_emit_canonical_a1() -> None:
    assert parse_cell_ref("  $aa$42 ") == (42, 27)
    assert make_cell_ref(1_048_576, 16_384) == "XFD1048576"


@pytest.mark.parametrize("ref", ["A0", "1A", "A", "A1048577", "Sheet1!A1"])
def test_cell_reference_parser_rejects_invalid_or_out_of_bounds_refs(ref: str) -> None:
    with pytest.raises(ValueError):
        parse_cell_ref(ref)


@pytest.mark.parametrize("row", [0, -1, 1_048_577])
def test_make_cell_ref_rejects_out_of_bounds_rows(row: int) -> None:
    with pytest.raises(ValueError):
        make_cell_ref(row, 1)


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("B2", Rect(2, 2, 2, 2)),
        ("$D$9:$B$2", Rect(2, 9, 2, 4)),
        ("$D:$B", Rect(1, 1_048_576, 2, 4)),
        ("$9:$2", Rect(2, 9, 1, 16_384)),
    ],
)
def test_parse_rect_supports_cells_and_reversed_whole_axis_ranges(
    source: str, expected: Rect
) -> None:
    assert parse_rect(source) == expected


@pytest.mark.parametrize(
    "source",
    ["", ":A1", "A1:", "A1:B2:C3", "A", "1", "A1:B", "XFE:XFE"],
)
def test_parse_rect_rejects_malformed_or_out_of_bounds_ranges(source: str) -> None:
    with pytest.raises(ValueError):
        parse_rect(source)


def test_contains_uses_inclusive_rectangle_boundaries() -> None:
    rect = Rect(2, 5, 3, 7)
    assert contains(rect, 2, 3)
    assert contains(rect, 5, 7)
    assert not contains(rect, 1, 3)
    assert not contains(rect, 5, 8)


@pytest.mark.parametrize(
    "format_code",
    ["yyyy-mm-dd", "[$-409]mmm-yy", "[Red]h:mm", "0.00 s"],
)
def test_custom_number_format_detects_date_tokens_outside_ignored_sections(
    format_code: str,
) -> None:
    assert custom_format_is_date(format_code)


@pytest.mark.parametrize(
    "format_code",
    ["0.00", '0 "days"', "0\\d", "0_d", "0*m", "[Red]0.00", "[<=1][Blue]0"],
)
def test_custom_number_format_ignores_quotes_brackets_and_escaped_tokens(
    format_code: str,
) -> None:
    assert not custom_format_is_date(format_code)


@pytest.mark.parametrize("num_fmt_id", sorted(BUILTIN_DATE_FORMAT_IDS))
def test_every_frozen_builtin_date_format_id_is_classified(num_fmt_id: int) -> None:
    root = etree.fromstring(
        f'<styleSheet><cellXfs><xf numFmtId="{num_fmt_id}"/></cellXfs></styleSheet>'
    )
    assert parse_style_catalog(root).is_date_style(0)


@pytest.mark.parametrize("num_fmt_id", [0, 13, 23, 26, 37, 44, 48, 49, 59, 164])
def test_adjacent_nondate_builtin_ids_are_not_classified(num_fmt_id: int) -> None:
    root = etree.fromstring(
        f'<styleSheet><cellXfs><xf numFmtId="{num_fmt_id}"/></cellXfs></styleSheet>'
    )
    assert not parse_style_catalog(root).is_date_style(0)


def test_style_catalog_parses_fonts_fills_colors_and_date_formats() -> None:
    root = etree.fromstring(
        b"""
        <styleSheet xmlns="urn:test-spreadsheetml">
          <!-- comments are not style sections -->
          <numFmts>
            <numFmt numFmtId="164" formatCode="0.00"/>
            <numFmt numFmtId="165" formatCode="yyyy-mm-dd"/>
          </numFmts>
          <fonts>
            <font/>
            <font><b/></font>
            <font><b val="0"/></font>
          </fonts>
          <fills>
            <fill/>
            <fill><patternFill patternType="none"/></fill>
            <fill>
              <patternFill patternType="solid">
                <fgColor rgb="FFFF0000"/><bgColor indexed="8"/>
              </patternFill>
            </fill>
            <fill>
              <patternFill patternType="solid">
                <fgColor theme="4"/><bgColor auto="1"/>
              </patternFill>
            </fill>
          </fills>
          <cellXfs>
            <xf/>
            <xf numFmtId="14" fontId="1" fillId="2"/>
            <xf numFmtId="164" fontId="0" fillId="0"/>
            <xf numFmtId="165" fontId="2" fillId="3"/>
          </cellXfs>
        </styleSheet>
        """
    )

    catalog = parse_style_catalog(root)

    assert catalog.custom_num_formats == {164: "0.00", 165: "yyyy-mm-dd"}
    assert [font.bold for font in catalog.fonts] == [False, True, False]
    assert catalog.fills[0].pattern_type is None
    assert catalog.fills[1].foreground is None
    assert (catalog.fills[2].foreground, catalog.fills[2].background) == (
        "rgb:FFFF0000",
        "indexed:8",
    )
    assert (catalog.fills[3].foreground, catalog.fills[3].background) == (
        "theme:4",
        "auto:1",
    )
    assert catalog.cell_xfs == (
        CellStyle(0, 0, 0, False),
        CellStyle(14, 1, 2, True),
        CellStyle(164, 0, 0, False),
        CellStyle(165, 2, 3, True),
    )
    assert catalog.is_date_style(1)
    assert catalog.is_date_style(3)
    assert not catalog.is_date_style(-1)
    assert not catalog.is_date_style(99)


def test_empty_style_catalog_supplies_safe_defaults() -> None:
    root = etree.fromstring(b"<styleSheet><!-- ignored --><unknown/></styleSheet>")

    catalog = parse_style_catalog(root)

    assert catalog.cell_xfs == (CellStyle(0, 0, 0, False),)
    assert [font.bold for font in catalog.fonts] == [False]
    assert catalog.fills[0].pattern_type is None
    assert catalog.custom_num_formats == {}


@pytest.mark.parametrize(
    ("num_fmt", "message"),
    [
        ('<numFmt formatCode="0"/>', "numFmt is missing numFmtId"),
        ('<numFmt numFmtId="164"/>', "numFmt is missing formatCode"),
    ],
)
def test_style_catalog_rejects_incomplete_custom_number_formats(num_fmt: str, message: str) -> None:
    root = etree.fromstring(f"<styleSheet><numFmts>{num_fmt}</numFmts></styleSheet>")

    with pytest.raises(ValueError, match=message):
        parse_style_catalog(root)
