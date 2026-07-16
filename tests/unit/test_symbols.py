"""Frozen symbol-id and normalized-header contracts."""

from __future__ import annotations

import pytest

from excel_lsp.core.symbols import (
    cell_symbol_id,
    column_symbol_id,
    deduplicate_normalized_headers,
    defined_name_symbol_id,
    formula_block_symbol_id,
    normalize_header,
    region_symbol_id,
    sheet_symbol_id,
)


def test_frozen_symbol_ids_match_the_handoff_scheme() -> None:
    assert sheet_symbol_id("My Sheet") == "sheet:My Sheet"
    assert region_symbol_id("My Sheet", 2) == "region:My Sheet:2"
    assert column_symbol_id("My Sheet", 2, "net_revenue#2") == ("col:My Sheet:2:net_revenue#2")
    assert defined_name_symbol_id("TaxRate") == "name:TaxRate"
    assert defined_name_symbol_id("TaxRate", scope_sheet="Inputs") == ("name:Inputs!TaxRate")
    assert formula_block_symbol_id("Calc", 4) == "fblock:Calc:4"
    assert cell_symbol_id("Résumé", "C10") == "cell:Résumé!C10"


def test_headers_normalize_unicode_spacing_and_duplicate_suffixes() -> None:
    assert normalize_header("  Net   Revenue ($) ", fallback="Column A") == "net_revenue"
    assert normalize_header("\uff32\uff21\uff34\uff25", fallback="Column B") == "rate"
    assert deduplicate_normalized_headers(
        (("Revenue", "Column A"), ("revenue!", "Column B"), ("", "Column C"))
    ) == ("revenue", "revenue#2", "column_c")


@pytest.mark.parametrize(
    ("call", "message"),
    (
        (lambda: region_symbol_id("Data", -1), "region ordinal"),
        (lambda: formula_block_symbol_id("Data", -1), "formula block ordinal"),
        (lambda: column_symbol_id("Data", 0, ""), "normalized header"),
        (lambda: normalize_header("", fallback=""), "fallback"),
    ),
)
def test_symbol_helpers_reject_ambiguous_components(call: object, message: str) -> None:
    assert callable(call)
    with pytest.raises(ValueError, match=message):
        call()
