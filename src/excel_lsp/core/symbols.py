"""Frozen public symbol identifiers and deterministic header normalization."""

from __future__ import annotations

import unicodedata
from collections.abc import Iterable


def sheet_symbol_id(sheet_name: str) -> str:
    """Return the frozen sheet symbol id."""
    return f"sheet:{sheet_name}"


def region_symbol_id(sheet_name: str, region_n: int) -> str:
    """Return the frozen row/column-ordered region symbol id."""
    _require_ordinal(region_n, "region")
    return f"region:{sheet_name}:{region_n}"


def column_symbol_id(sheet_name: str, region_n: int, normalized_header: str) -> str:
    """Return the frozen column symbol id for one unique normalized header."""
    _require_ordinal(region_n, "region")
    if not normalized_header:
        raise ValueError("normalized header must not be empty")
    return f"col:{sheet_name}:{region_n}:{normalized_header}"


def defined_name_symbol_id(name: str, *, scope_sheet: str | None = None) -> str:
    """Return the frozen workbook- or sheet-scoped defined-name id."""
    if scope_sheet is None:
        return f"name:{name}"
    return f"name:{scope_sheet}!{name}"


def formula_block_symbol_id(sheet_name: str, block_n: int) -> str:
    """Return the frozen formula-block symbol id."""
    _require_ordinal(block_n, "formula block")
    return f"fblock:{sheet_name}:{block_n}"


def cell_symbol_id(sheet_name: str, ref: str) -> str:
    """Return the frozen cell symbol id."""
    return f"cell:{sheet_name}!{ref}"


def normalize_header(header: str, *, fallback: str) -> str:
    """Normalize a display header into a compact, Unicode-stable id component."""
    normalized = unicodedata.normalize("NFKC", header).casefold().strip()
    pieces: list[str] = []
    separator_pending = False
    for character in normalized:
        if character.isalnum():
            if separator_pending and pieces:
                pieces.append("_")
            pieces.append(character)
            separator_pending = False
        else:
            separator_pending = True
    value = "".join(pieces).strip("_")
    if value:
        return value
    fallback_value = unicodedata.normalize("NFKC", fallback).casefold().strip()
    if not fallback_value:
        raise ValueError("header fallback must not normalize to empty")
    return normalize_header(fallback_value, fallback="column")


def deduplicate_normalized_headers(headers: Iterable[tuple[str, str]]) -> tuple[str, ...]:
    """Normalize display/fallback pairs and suffix duplicates with frozen ``#k``."""
    counts: dict[str, int] = {}
    result: list[str] = []
    for header, fallback in headers:
        base = normalize_header(header, fallback=fallback)
        count = counts.get(base, 0) + 1
        counts[base] = count
        result.append(base if count == 1 else f"{base}#{count}")
    return tuple(result)


def _require_ordinal(value: int, label: str) -> None:
    if value < 0:
        raise ValueError(f"{label} ordinal must be nonnegative")


__all__ = [
    "cell_symbol_id",
    "column_symbol_id",
    "deduplicate_normalized_headers",
    "defined_name_symbol_id",
    "formula_block_symbol_id",
    "normalize_header",
    "region_symbol_id",
    "sheet_symbol_id",
]
