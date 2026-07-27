"""A1 reference geometry shared by formula normalization and extraction.

The parser in this module is deliberately lexical.  It preserves a workbook,
worksheet, 3-D, or external qualifier verbatim and only interprets the local
cell, row, or column coordinates.  Resolution of the qualifier belongs to the
higher-level reference analyzer.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from excel_lsp.core.models import Rect

_MAX_ROW = 1_048_576
_MAX_COLUMN = 16_384

_CELL_RE = re.compile(
    r"(?P<col_abs>\$?)(?P<col>[A-Za-z]{1,3})"
    r"(?P<row_abs>\$?)(?P<row>[1-9][0-9]{0,6})"
)
_COLUMN_RE = re.compile(r"(?P<absolute>\$?)(?P<column>[A-Za-z]{1,3})")
_ROW_RE = re.compile(r"(?P<absolute>\$?)(?P<row>[1-9][0-9]{0,6})")


@dataclass(frozen=True, slots=True)
class CellRef:
    """One local, one-based worksheet coordinate."""

    row: int
    col: int

    def __post_init__(self) -> None:
        if type(self.row) is not int or type(self.col) is not int:
            raise ValueError("cell coordinates must be integers")
        if not 1 <= self.row <= _MAX_ROW:
            raise ValueError("cell row exceeds Excel bounds")
        if not 1 <= self.col <= _MAX_COLUMN:
            raise ValueError("cell column exceeds Excel bounds")


@dataclass(frozen=True, slots=True)
class AxisTerm:
    """One absolute coordinate or signed coordinate-relative displacement."""

    relative: bool
    value: int

    def __post_init__(self) -> None:
        if type(self.relative) is not bool:
            raise ValueError("axis relative flag must be a boolean")
        if type(self.value) is not int:
            raise ValueError("axis term value must be an integer")
        if not self.relative and self.value < 1:
            raise ValueError("absolute axis coordinates must be positive")


@dataclass(frozen=True, slots=True)
class ReferenceGeometry:
    """Two row endpoints and two column endpoints for one A1 area."""

    row_a: AxisTerm
    row_b: AxisTerm
    col_a: AxisTerm
    col_b: AxisTerm


@dataclass(frozen=True, slots=True)
class ParsedA1Reference:
    """A local area plus its uninterpreted, verbatim qualifier."""

    qualifier: str
    geometry: ReferenceGeometry


@dataclass(frozen=True, slots=True)
class ParsedModernA1Range:
    """One A1 range whose endpoints carry ``@`` and/or ``#`` operators."""

    parsed: ParsedA1Reference
    left_original: str
    right_original: str
    left_implicit: bool
    left_spill: bool
    right_implicit: bool
    right_spill: bool


@dataclass(frozen=True, slots=True)
class ParsedRangeEndpoint:
    """One side of a range operator, with modern operators retained."""

    original: str
    core: str
    implicit: bool
    spill: bool
    parsed_a1: ParsedA1Reference | None


@dataclass(frozen=True, slots=True)
class ParsedReferenceRange:
    """A range operator with at least one non-A1 reference endpoint."""

    left: ParsedRangeEndpoint
    right: ParsedRangeEndpoint


# Backward-compatible descriptive alias for the original one-A1/one-name
# surface.  New callers should use ``ParsedReferenceRange``.
ParsedMixedA1Range = ParsedReferenceRange


def parse_a1_reference(text: str, anchor: CellRef) -> ParsedA1Reference | None:
    """Parse a bounded A1 cell/range while preserving its qualifier.

    ``None`` means the operand is not an A1 reference handled here.  In
    particular, defined names, structured references, spill references, and
    R1C1-looking identifiers remain available to the caller unchanged.
    """
    if not text or text.endswith("#") or text.startswith("@"):
        return None
    split = _split_qualifier(text)
    if split is None:
        return None
    qualifier, local = split
    pieces = local.split(":")
    if len(pieces) not in {1, 2} or any(not piece for piece in pieces):
        return None

    cells = tuple(_parse_cell(piece, anchor) for piece in pieces)
    if all(cell is not None for cell in cells):
        first = cells[0]
        second = cells[-1]
        assert first is not None and second is not None
        return ParsedA1Reference(
            qualifier,
            ReferenceGeometry(first[0], second[0], first[1], second[1]),
        )

    if len(pieces) == 2:
        columns = tuple(_parse_column(piece, anchor) for piece in pieces)
        if all(column is not None for column in columns):
            first_col = columns[0]
            second_col = columns[1]
            assert first_col is not None and second_col is not None
            return ParsedA1Reference(
                qualifier,
                ReferenceGeometry(
                    AxisTerm(False, 1),
                    AxisTerm(False, _MAX_ROW),
                    first_col,
                    second_col,
                ),
            )

        rows = tuple(_parse_row(piece, anchor) for piece in pieces)
        if all(row is not None for row in rows):
            first_row = rows[0]
            second_row = rows[1]
            assert first_row is not None and second_row is not None
            return ParsedA1Reference(
                qualifier,
                ReferenceGeometry(
                    first_row,
                    second_row,
                    AxisTerm(False, 1),
                    AxisTerm(False, _MAX_COLUMN),
                ),
            )
    return None


def parse_modern_a1_range(text: str, anchor: CellRef) -> ParsedModernA1Range | None:
    """Parse modern operators attached to either endpoint of one A1 range.

    ``openpyxl`` returns constructs such as ``A1#:B5`` and ``A1:@B5`` as one
    RANGE operand.  Removing operators only from the complete operand loses
    their endpoint binding.  This helper locates the local range colon while
    ignoring quoted 3-D qualifiers and bracketed external-workbook text, then
    parses the operator-free A1 range through the ordinary geometry layer.
    """
    colon = _local_range_colon(text)
    if colon is None:
        return None
    left_original = text[:colon]
    right_original = text[colon + 1 :]
    left = _modern_endpoint(left_original)
    right = _modern_endpoint(right_original)
    if left is None or right is None:
        return None
    left_core, left_implicit, left_spill = left
    right_core, right_implicit, right_spill = right
    if not (left_implicit or left_spill or right_implicit or right_spill):
        return None
    parsed = parse_a1_reference(f"{left_core}:{right_core}", anchor)
    if parsed is None:
        return None
    return ParsedModernA1Range(
        parsed=parsed,
        left_original=left_original,
        right_original=right_original,
        left_implicit=left_implicit,
        left_spill=left_spill,
        right_implicit=right_implicit,
        right_spill=right_spill,
    )


def parse_mixed_a1_range(text: str, anchor: CellRef) -> ParsedMixedA1Range | None:
    """Parse a range with one A1 endpoint and one non-A1 endpoint.

    Excel permits the range operator to join a defined name and a cell, for
    example ``Rate:B5`` or ``A1:Rate``.  The ordinary A1 parser must not absorb
    the name, while formula translation must still move the A1 side.  Colon
    selection ignores quoted 3-D spans and bracketed external-workbook text.
    The conservative fallback before a qualifier only accepts a quoted or
    bracketed right endpoint, avoiding the ambiguous unquoted ``Jan:Mar!A1``
    spelling.
    """
    parsed = parse_reference_range(text, anchor)
    if parsed is None:
        return None
    if (parsed.left.parsed_a1 is None) == (parsed.right.parsed_a1 is None):
        return None
    return parsed


def parse_reference_range(text: str, anchor: CellRef) -> ParsedReferenceRange | None:
    """Choose a context-free range split only when it is lexically safe."""
    candidates = parse_reference_range_candidates(text, anchor)
    if not candidates:
        return None
    scores = tuple(_context_free_range_score(candidate) for candidate in candidates)
    best = max(scores)
    winners = [
        candidate for candidate, score in zip(candidates, scores, strict=True) if score == best
    ]
    return winners[0] if len(winners) == 1 else None


def parse_reference_range_candidates(
    text: str,
    anchor: CellRef,
) -> tuple[ParsedReferenceRange, ...]:
    """Enumerate quote/bracket/3-D-safe semantic splits for one colon chain."""
    whole_a1 = parse_a1_reference(text, anchor)
    candidates: list[ParsedReferenceRange] = []
    for colon in _range_operator_colons(text):
        left = _parse_range_endpoint(text[:colon], anchor)
        right = _parse_range_endpoint(text[colon + 1 :], anchor)
        if left is None or right is None:
            continue
        if any(
            endpoint.parsed_a1 is None and _top_level_colons(endpoint.core)
            for endpoint in (left, right)
        ):
            continue
        qualified_right = (
            left.parsed_a1 is None
            and right.parsed_a1 is not None
            and right.core.startswith(("'", "["))
        )
        # Whole-column, whole-row, and ordinary cell ranges take precedence
        # even though their individual endpoints are not standalone areas.
        # The qualified-right exception prevents ``Name:'Sheet'!A1`` from
        # being swallowed as one malformed qualifier.
        if whole_a1 is not None and not qualified_right:
            continue
        candidate = ParsedReferenceRange(left, right)
        if candidate not in candidates:
            candidates.append(candidate)
    return tuple(candidates)


def _context_free_range_score(parsed: ParsedReferenceRange) -> tuple[int, int]:
    endpoints = (parsed.left, parsed.right)
    return (
        sum(endpoint.parsed_a1 is not None for endpoint in endpoints),
        sum(
            _is_repeated_whole_axis(endpoint.core)
            for endpoint in endpoints
            if endpoint.parsed_a1 is not None
        ),
    )


def decorate_modern_a1_range(
    rendered: str,
    modern: ParsedModernA1Range,
    *,
    preserve_spill_endpoints: bool = False,
) -> str:
    """Restore endpoint operators onto a rendered A1 or R1C1 range."""
    colon = _local_range_colon(rendered)
    if colon is None:  # pragma: no cover - caller rendered ``modern.parsed``
        raise ValueError("rendered modern range has no local range operator")
    left = rendered[:colon]
    right = rendered[colon + 1 :]
    if preserve_spill_endpoints and modern.left_spill:
        left = modern.left_original
    else:
        left = _decorate_modern_endpoint(
            left,
            implicit=modern.left_implicit,
            spill=modern.left_spill,
        )
    if preserve_spill_endpoints and modern.right_spill:
        right = modern.right_original
    else:
        right = _decorate_modern_endpoint(
            right,
            implicit=modern.right_implicit,
            spill=modern.right_spill,
        )
    return f"{left}:{right}"


def modern_range_endpoint_geometries(
    modern: ParsedModernA1Range,
) -> tuple[ReferenceGeometry, ReferenceGeometry]:
    """Return the left and right endpoint geometries of a modern range."""
    geometry = modern.parsed.geometry
    whole_rows = _is_absolute_span(geometry.row_a, geometry.row_b, 1, _MAX_ROW)
    whole_cols = _is_absolute_span(geometry.col_a, geometry.col_b, 1, _MAX_COLUMN)
    if whole_rows and not whole_cols:
        left = ReferenceGeometry(
            geometry.row_a,
            geometry.row_b,
            geometry.col_a,
            geometry.col_a,
        )
        right = ReferenceGeometry(
            geometry.row_a,
            geometry.row_b,
            geometry.col_b,
            geometry.col_b,
        )
    elif whole_cols and not whole_rows:
        left = ReferenceGeometry(
            geometry.row_a,
            geometry.row_a,
            geometry.col_a,
            geometry.col_b,
        )
        right = ReferenceGeometry(
            geometry.row_b,
            geometry.row_b,
            geometry.col_a,
            geometry.col_b,
        )
    else:
        left = ReferenceGeometry(
            geometry.row_a,
            geometry.row_a,
            geometry.col_a,
            geometry.col_a,
        )
        right = ReferenceGeometry(
            geometry.row_b,
            geometry.row_b,
            geometry.col_b,
            geometry.col_b,
        )
    return left, right


def resolve_reference(geometry: ReferenceGeometry, anchor: CellRef) -> Rect:
    """Resolve one geometry at a concrete formula cell."""
    row_a = _resolve_term(geometry.row_a, anchor.row)
    row_b = _resolve_term(geometry.row_b, anchor.row)
    col_a = _resolve_term(geometry.col_a, anchor.col)
    col_b = _resolve_term(geometry.col_b, anchor.col)
    return Rect(
        min(row_a, row_b),
        max(row_a, row_b),
        min(col_a, col_b),
        max(col_a, col_b),
    )


def extrude_reference(geometry: ReferenceGeometry, source_block: Rect) -> Rect:
    """Return the exact in-bounds union of an area over a source block."""
    rows = (
        *_term_extrema(geometry.row_a, source_block.row_min, source_block.row_max),
        *_term_extrema(geometry.row_b, source_block.row_min, source_block.row_max),
    )
    cols = (
        *_term_extrema(geometry.col_a, source_block.col_min, source_block.col_max),
        *_term_extrema(geometry.col_b, source_block.col_min, source_block.col_max),
    )
    row_min = max(1, min(rows))
    row_max = min(_MAX_ROW, max(rows))
    col_min = max(1, min(cols))
    col_max = min(_MAX_COLUMN, max(cols))
    if row_min > row_max or col_min > col_max:
        raise ValueError("extruded reference does not intersect the worksheet")
    return Rect(row_min, row_max, col_min, col_max)


def render_r1c1_reference(
    parsed: ParsedA1Reference,
    *,
    preserve_range: bool = False,
) -> str:
    """Render parsed geometry into one canonical R1C1 reference token."""
    geometry = parsed.geometry
    whole_rows = _is_absolute_span(geometry.row_a, geometry.row_b, 1, _MAX_ROW)
    whole_cols = _is_absolute_span(geometry.col_a, geometry.col_b, 1, _MAX_COLUMN)
    if whole_rows and not whole_cols:
        local = f"C{_render_axis_term(geometry.col_a)}:C{_render_axis_term(geometry.col_b)}"
    elif whole_cols and not whole_rows:
        local = f"R{_render_axis_term(geometry.row_a)}:R{_render_axis_term(geometry.row_b)}"
    else:
        first = f"R{_render_axis_term(geometry.row_a)}C{_render_axis_term(geometry.col_a)}"
        second = f"R{_render_axis_term(geometry.row_b)}C{_render_axis_term(geometry.col_b)}"
        local = f"{first}:{second}" if preserve_range or first != second else first
    return f"{parsed.qualifier}{local}"


def render_a1_reference(
    parsed: ParsedA1Reference,
    anchor: CellRef,
    *,
    preserve_range: bool = False,
) -> str:
    """Render parsed geometry as a bounded A1 reference at a new anchor."""
    geometry = parsed.geometry
    whole_rows = _is_absolute_span(geometry.row_a, geometry.row_b, 1, _MAX_ROW)
    whole_cols = _is_absolute_span(geometry.col_a, geometry.col_b, 1, _MAX_COLUMN)
    if whole_rows and not whole_cols:
        first = _render_a1_axis(geometry.col_a, anchor.col, column=True)
        second = _render_a1_axis(geometry.col_b, anchor.col, column=True)
        local = f"{first}:{second}"
    elif whole_cols and not whole_rows:
        first = _render_a1_axis(geometry.row_a, anchor.row, column=False)
        second = _render_a1_axis(geometry.row_b, anchor.row, column=False)
        local = f"{first}:{second}"
    else:
        first = (
            f"{_render_a1_axis(geometry.col_a, anchor.col, column=True)}"
            f"{_render_a1_axis(geometry.row_a, anchor.row, column=False)}"
        )
        second = (
            f"{_render_a1_axis(geometry.col_b, anchor.col, column=True)}"
            f"{_render_a1_axis(geometry.row_b, anchor.row, column=False)}"
        )
        local = f"{first}:{second}" if preserve_range or first != second else first
    return f"{parsed.qualifier}{local}"


def _split_qualifier(text: str) -> tuple[str, str] | None:
    bangs: list[int] = []
    in_quote = False
    bracket_depth = 0
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
                    return None
            elif character == "!" and bracket_depth == 0:
                bangs.append(index)
        index += 1
    if in_quote or bracket_depth != 0 or len(bangs) > 1:
        return None
    if not bangs:
        return "", text
    bang = bangs[0]
    if bang == 0 or bang == len(text) - 1:
        return None
    return text[: bang + 1], text[bang + 1 :]


def _local_range_colon(text: str) -> int | None:
    """Locate one range colon in the local coordinate portion of ``text``."""
    colons: list[int] = []
    bangs: list[int] = []
    in_quote = False
    bracket_depth = 0
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
                    return None
            elif bracket_depth == 0:
                if character == "!":
                    bangs.append(index)
                elif character == ":":
                    colons.append(index)
        index += 1
    if in_quote or bracket_depth != 0:
        return None
    qualifier_end = bangs[-1] if bangs else -1
    local_colons = [colon for colon in colons if colon > qualifier_end]
    return local_colons[0] if len(local_colons) == 1 else None


def _range_operator_colons(text: str) -> tuple[int, ...]:
    """Return safe outer-range candidates, excluding qualifier colons."""
    colons: list[int] = []
    bangs: list[int] = []
    in_quote = False
    bracket_depth = 0
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
                    return ()
            elif bracket_depth == 0:
                if character == ":":
                    colons.append(index)
                elif character == "!":
                    bangs.append(index)
        index += 1
    if in_quote or bracket_depth != 0 or len(bangs) > 1:
        return ()
    qualifier_end = bangs[-1] if bangs else -1
    result: list[int] = []
    for colon in colons:
        if colon > qualifier_end:
            result.append(colon)
            continue
        # ``Name:'Quoted Sheet'!A1`` and ``Name:[1]Sheet!A1`` put the range
        # operator before the right qualifier. This is distinguishable from
        # an unquoted 3-D qualifier such as ``Jan:Mar!A1``.
        right = _modern_endpoint(text[colon + 1 :])
        if right is not None and right[0].startswith(("'", "[")):
            result.append(colon)
    return tuple(result)


def _top_level_colons(text: str) -> tuple[int, ...]:
    """Find colons outside quotes and brackets in one endpoint candidate."""
    colons: list[int] = []
    in_quote = False
    bracket_depth = 0
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
                    return ()
            elif character == ":" and bracket_depth == 0:
                colons.append(index)
        index += 1
    if in_quote or bracket_depth != 0:
        return ()
    return tuple(colons)


def _is_repeated_whole_axis(text: str) -> bool:
    split = _split_qualifier(text)
    if split is None:
        return False
    _qualifier, local = split
    pieces = local.split(":")
    if len(pieces) != 2:
        return False
    left, right = (piece.replace("$", "").casefold() for piece in pieces)
    if left != right:
        return False
    return _COLUMN_RE.fullmatch(pieces[0]) is not None or _ROW_RE.fullmatch(pieces[0]) is not None


def _parse_range_endpoint(text: str, anchor: CellRef) -> ParsedRangeEndpoint | None:
    modern = _modern_endpoint(text)
    if modern is None:
        return None
    core, implicit, spill = modern
    return ParsedRangeEndpoint(
        original=text,
        core=core,
        implicit=implicit,
        spill=spill,
        parsed_a1=parse_a1_reference(core, anchor),
    )


def _modern_endpoint(text: str) -> tuple[str, bool, bool] | None:
    implicit = text.startswith("@")
    core = text[1:] if implicit else text
    spill = core.endswith("#")
    if spill:
        core = core[:-1]
    if not core:
        return None
    return core, implicit, spill


def _decorate_modern_endpoint(text: str, *, implicit: bool, spill: bool) -> str:
    return f"{'@' if implicit else ''}{text}{'#' if spill else ''}"


def _parse_cell(text: str, anchor: CellRef) -> tuple[AxisTerm, AxisTerm] | None:
    match = _CELL_RE.fullmatch(text)
    if match is None:
        return None
    try:
        column = _column_number(match.group("col"))
    except ValueError:
        return None
    row = int(match.group("row"))
    if row > _MAX_ROW:
        return None
    row_absolute = bool(match.group("row_abs"))
    col_absolute = bool(match.group("col_abs"))
    return (
        AxisTerm(not row_absolute, row if row_absolute else row - anchor.row),
        AxisTerm(not col_absolute, column if col_absolute else column - anchor.col),
    )


def _parse_column(text: str, anchor: CellRef) -> AxisTerm | None:
    match = _COLUMN_RE.fullmatch(text)
    if match is None:
        return None
    try:
        column = _column_number(match.group("column"))
    except ValueError:
        return None
    absolute = bool(match.group("absolute"))
    return AxisTerm(not absolute, column if absolute else column - anchor.col)


def _parse_row(text: str, anchor: CellRef) -> AxisTerm | None:
    match = _ROW_RE.fullmatch(text)
    if match is None:
        return None
    row = int(match.group("row"))
    if row > _MAX_ROW:
        return None
    absolute = bool(match.group("absolute"))
    return AxisTerm(not absolute, row if absolute else row - anchor.row)


def _resolve_term(term: AxisTerm, anchor: int) -> int:
    return anchor + term.value if term.relative else term.value


def _term_extrema(term: AxisTerm, lower: int, upper: int) -> tuple[int, int]:
    if term.relative:
        return lower + term.value, upper + term.value
    return term.value, term.value


def _render_axis_term(term: AxisTerm) -> str:
    if not term.relative:
        return str(term.value)
    if term.value == 0:
        return ""
    sign = "+" if term.value > 0 else ""
    return f"[{sign}{term.value}]"


def _render_a1_axis(term: AxisTerm, anchor: int, *, column: bool) -> str:
    resolved = _resolve_term(term, anchor)
    maximum = _MAX_COLUMN if column else _MAX_ROW
    if not 1 <= resolved <= maximum:
        raise ValueError("translated reference exceeds Excel bounds")
    rendered = _column_label(resolved) if column else str(resolved)
    return f"${rendered}" if not term.relative else rendered


def _column_number(label: str) -> int:
    result = 0
    for character in label.upper():
        if not "A" <= character <= "Z":
            raise ValueError(f"invalid column label: {label!r}")
        result = result * 26 + ord(character) - ord("A") + 1
    if not 1 <= result <= _MAX_COLUMN:
        raise ValueError(f"column label exceeds Excel bounds: {label!r}")
    return result


def _column_label(column: int) -> str:
    if not 1 <= column <= _MAX_COLUMN:
        raise ValueError(f"column number exceeds Excel bounds: {column}")
    result = ""
    value = column
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _is_absolute_span(
    first: AxisTerm,
    second: AxisTerm,
    lower: int,
    upper: int,
) -> bool:
    return (
        not first.relative and not second.relative and {first.value, second.value} == {lower, upper}
    )


__all__ = [
    "AxisTerm",
    "CellRef",
    "ParsedA1Reference",
    "ParsedMixedA1Range",
    "ParsedModernA1Range",
    "ParsedRangeEndpoint",
    "ParsedReferenceRange",
    "ReferenceGeometry",
    "decorate_modern_a1_range",
    "extrude_reference",
    "modern_range_endpoint_geometries",
    "parse_a1_reference",
    "parse_mixed_a1_range",
    "parse_modern_a1_range",
    "parse_reference_range",
    "parse_reference_range_candidates",
    "render_a1_reference",
    "render_r1c1_reference",
    "resolve_reference",
]
