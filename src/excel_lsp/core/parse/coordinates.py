"""Minimal A1 coordinate handling used by the OOXML reader."""

from __future__ import annotations

import re

from excel_lsp.core.models import Rect

_CELL_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})")
_COLUMN_RE = re.compile(r"\$?([A-Za-z]{1,3})")
_ROW_RE = re.compile(r"\$?([1-9][0-9]{0,6})")


def column_number(label: str) -> int:
    """Convert an Excel column label to a one-based number."""
    value = 0
    for character in label.upper():
        if not "A" <= character <= "Z":
            raise ValueError(f"invalid column label: {label!r}")
        value = value * 26 + ord(character) - ord("A") + 1
    if not 1 <= value <= 16_384:
        raise ValueError(f"column exceeds Excel bounds: {label!r}")
    return value


def column_label(column: int) -> str:
    """Convert a one-based number to an Excel column label."""
    if not 1 <= column <= 16_384:
        raise ValueError(f"column exceeds Excel bounds: {column}")
    result: list[str] = []
    while column:
        column, remainder = divmod(column - 1, 26)
        result.append(chr(ord("A") + remainder))
    return "".join(reversed(result))


def parse_cell_ref(ref: str) -> tuple[int, int]:
    """Return ``(row, column)`` for one local A1 cell reference."""
    match = _CELL_RE.fullmatch(ref.strip())
    if match is None:
        raise ValueError(f"invalid cell reference: {ref!r}")
    row = int(match.group(2))
    if row > 1_048_576:
        raise ValueError(f"row exceeds Excel bounds: {ref!r}")
    return row, column_number(match.group(1))


def make_cell_ref(row: int, column: int) -> str:
    """Create a canonical local A1 reference."""
    if not 1 <= row <= 1_048_576:
        raise ValueError(f"row exceeds Excel bounds: {row}")
    return f"{column_label(column)}{row}"


def parse_rect(ref: str) -> Rect:
    """Parse a cell, rectangular, whole-column, or whole-row local reference."""
    value = ref.strip()
    pieces = value.split(":")
    if len(pieces) > 2 or not pieces[0] or (len(pieces) == 2 and not pieces[1]):
        raise ValueError(f"invalid rectangular reference: {ref!r}")
    left = pieces[0]
    right = pieces[-1]

    left_cell = _CELL_RE.fullmatch(left)
    right_cell = _CELL_RE.fullmatch(right)
    if left_cell is not None and right_cell is not None:
        row_a, col_a = parse_cell_ref(left)
        row_b, col_b = parse_cell_ref(right)
        return Rect(min(row_a, row_b), max(row_a, row_b), min(col_a, col_b), max(col_a, col_b))

    left_col = _COLUMN_RE.fullmatch(left)
    right_col = _COLUMN_RE.fullmatch(right)
    if len(pieces) == 2 and left_col is not None and right_col is not None:
        col_a = column_number(left_col.group(1))
        col_b = column_number(right_col.group(1))
        return Rect(1, 1_048_576, min(col_a, col_b), max(col_a, col_b))

    left_row = _ROW_RE.fullmatch(left)
    right_row = _ROW_RE.fullmatch(right)
    if len(pieces) == 2 and left_row is not None and right_row is not None:
        row_a = int(left_row.group(1))
        row_b = int(right_row.group(1))
        return Rect(min(row_a, row_b), max(row_a, row_b), 1, 16_384)

    raise ValueError(f"invalid rectangular reference: {ref!r}")


def contains(rect: Rect, row: int, column: int) -> bool:
    """Return whether a coordinate lies inside an inclusive rectangle."""
    return rect.row_min <= row <= rect.row_max and rect.col_min <= column <= rect.col_max
