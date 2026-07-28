"""Render an R1C1 formula at a concrete worksheet anchor."""

from __future__ import annotations

import re
from dataclasses import dataclass

from excel_lsp.core.formulas.a1 import CellRef, parse_a1_reference
from excel_lsp.core.formulas.tokens import FormulaToken, tokenize_formula
from excel_lsp.core.parse.coordinates import column_label

_AXIS = r"(?:\[(?P<{name}_rel>[+-]?\d+)\]|(?P<{name}_abs>[1-9]\d*)?)"
_CELL_LOCAL = re.compile(
    rf"R{_AXIS.format(name='row')}C{_AXIS.format(name='col')}",
    re.IGNORECASE,
)
_ROW_LOCAL = re.compile(rf"R{_AXIS.format(name='row')}", re.IGNORECASE)
_COL_LOCAL = re.compile(rf"C{_AXIS.format(name='col')}", re.IGNORECASE)


@dataclass(frozen=True, slots=True)
class _RenderedReference:
    text: str
    is_reference: bool


def from_r1c1(formula: str, anchor: CellRef) -> str | None:
    """Return A1 text when ``formula`` contains R1C1 refs, otherwise ``None``.

    A formula that mixes concrete A1 and R1C1 operands is rejected so a caller
    cannot silently fill a column under an ambiguous reference mode.
    """
    tokens = tokenize_formula(formula)
    rendered: list[str] = []
    saw_r1c1 = False
    saw_a1 = False
    for token in tokens:
        converted = _render_token(token, anchor)
        rendered.append(converted.text)
        saw_r1c1 = saw_r1c1 or converted.is_reference
        if (
            token.type == "OPERAND"
            and token.subtype == "RANGE"
            and not converted.is_reference
            and parse_a1_reference(token.value, anchor) is not None
        ):
            saw_a1 = True
    if not saw_r1c1:
        return None
    if saw_a1:
        raise ValueError("formula mixes A1 and R1C1 reference modes")
    return "=" + "".join(rendered)


def _render_token(token: FormulaToken, anchor: CellRef) -> _RenderedReference:
    if token.type == "FUNC" and token.subtype == "OPEN":
        body = token.value[:-1]
        if ":" in body:
            reference, function_name = body.rsplit(":", 1)
            converted = _render_reference(reference, anchor)
            if converted is not None:
                return _RenderedReference(f"{converted}:{function_name}(", True)
    if token.type == "OPERAND" and token.subtype == "RANGE":
        leading_colon = token.value.startswith(":")
        candidate = token.value[1:] if leading_colon else token.value
        converted = _render_reference(candidate, anchor)
        if converted is not None:
            prefix = ":" if leading_colon else ""
            return _RenderedReference(f"{prefix}{converted}", True)
    return _RenderedReference(token.value, False)


def _render_reference(text: str, anchor: CellRef) -> str | None:
    qualifier, local = _split_qualifier(text)
    if local is None:
        return None
    pieces = local.split(":")
    if len(pieces) not in {1, 2} or any(not piece for piece in pieces):
        return None

    cells = tuple(_render_cell(piece, anchor) for piece in pieces)
    if all(cell is not None for cell in cells):
        first = cells[0]
        second = cells[-1]
        assert first is not None and second is not None
        local_rendered = first if len(pieces) == 1 else f"{first}:{second}"
        return f"{qualifier}{local_rendered}"

    if len(pieces) == 2:
        rows = tuple(_render_row(piece, anchor.row) for piece in pieces)
        if all(row is not None for row in rows):
            return f"{qualifier}{rows[0]}:{rows[1]}"
        columns = tuple(_render_column(piece, anchor.col) for piece in pieces)
        if all(column is not None for column in columns):
            return f"{qualifier}{columns[0]}:{columns[1]}"
    return None


def _render_cell(text: str, anchor: CellRef) -> str | None:
    match = _CELL_LOCAL.fullmatch(text)
    if match is None:
        return None
    row = _axis_value(match, "row", anchor.row, 1_048_576, column=False)
    column = _axis_value(match, "col", anchor.col, 16_384, column=True)
    return f"{column}{row}"


def _render_row(text: str, anchor: int) -> str | None:
    match = _ROW_LOCAL.fullmatch(text)
    if match is None:
        return None
    return _axis_value(match, "row", anchor, 1_048_576, column=False)


def _render_column(text: str, anchor: int) -> str | None:
    match = _COL_LOCAL.fullmatch(text)
    if match is None:
        return None
    return _axis_value(match, "col", anchor, 16_384, column=True)


def _axis_value(
    match: re.Match[str],
    name: str,
    anchor: int,
    maximum: int,
    *,
    column: bool,
) -> str:
    relative = match.group(f"{name}_rel")
    absolute = match.group(f"{name}_abs")
    if relative is not None:
        value = anchor + int(relative)
        is_absolute = False
    elif absolute is not None:
        value = int(absolute)
        is_absolute = True
    else:
        value = anchor
        is_absolute = False
    if not 1 <= value <= maximum:
        raise ValueError("translated R1C1 reference exceeds worksheet bounds")
    rendered = column_label(value) if column else str(value)
    return f"${rendered}" if is_absolute else rendered


def _split_qualifier(text: str) -> tuple[str, str | None]:
    in_quote = False
    bracket_depth = 0
    bang = -1
    index = 0
    while index < len(text):
        character = text[index]
        if character == "'":
            if in_quote and index + 1 < len(text) and text[index + 1] == "'":
                index += 2
                continue
            in_quote = not in_quote
        elif not in_quote:
            if character == "[":
                bracket_depth += 1
            elif character == "]":
                bracket_depth -= 1
                if bracket_depth < 0:
                    return "", None
            elif character == "!" and bracket_depth == 0:
                if bang >= 0:
                    return "", None
                bang = index
        index += 1
    if in_quote or bracket_depth != 0:
        return "", None
    if bang < 0:
        return "", text
    if bang == 0 or bang == len(text) - 1:
        return "", None
    return text[: bang + 1], text[bang + 1 :]


__all__ = ["from_r1c1"]
