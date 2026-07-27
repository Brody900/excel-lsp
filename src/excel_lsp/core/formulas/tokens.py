"""Spill-aware immutable wrapper around openpyxl's formula tokenizer."""

from __future__ import annotations

from dataclasses import dataclass

from openpyxl.formula.tokenizer import Token, Tokenizer, TokenizerError


class FormulaSyntaxError(ValueError):
    """A formula cannot be tokenized or has unbalanced structural tokens."""


@dataclass(frozen=True, slots=True)
class FormulaToken:
    """One formula token with source spelling retained."""

    value: str
    type: str
    subtype: str


_OPEN_TYPES = frozenset({"FUNC", "PAREN", "ARRAY"})
WHITESPACE_TOKEN_TYPE = Token.WSPACE
_STRUCTURED_ESCAPES = {
    "'[": "ESC_LBRACKET",
    "']": "ESC_RBRACKET",
    "'#": "ESC_HASH",
    "'@": "ESC_AT",
    "''": "ESC_APOSTROPHE",
}


def tokenize_formula(formula: str) -> tuple[FormulaToken, ...]:
    """Tokenize a formula without losing modern spill operators or spaces.

    openpyxl 3.1.5 rejects postfix ``#``.  A quote/bracket-aware pass shields
    only genuine spill operators with a collision-free identifier suffix, then
    restores them in the immutable result.  The wrapper also rejects unbalanced
    structures that openpyxl otherwise accepts (notably ``=SUM(``).
    """
    source = formula if formula.startswith("=") else f"={formula}"
    if _supports_exact_fast_path(source):
        try:
            raw_tokens = Tokenizer(source).items
        except (TokenizerError, IndexError, ValueError) as error:
            raise FormulaSyntaxError(str(error)) from error
        exact = tuple(FormulaToken(token.value, token.type, token.subtype) for token in raw_tokens)
        if "".join(token.value for token in exact) == source[1:]:
            normalized = _normalize_intersection_groups(exact)
            _validate_structure(normalized)
            return normalized

    spill_marker = _unused_marker(source, "SPILL")
    implicit_qualifier_marker = _unused_marker(source, "IMPLICIT_QUALIFIER")
    structured_markers = {
        escape: _unused_marker(source, label) for escape, label in _STRUCTURED_ESCAPES.items()
    }
    shielded = _shield_modern_syntax(
        source,
        spill_marker,
        implicit_qualifier_marker,
        structured_markers,
    )
    tokenizer_source = _normalize_tokenizer_whitespace(shielded)
    try:
        raw_tokens = Tokenizer(tokenizer_source).items
    except (TokenizerError, IndexError, ValueError) as error:
        raise FormulaSyntaxError(str(error)) from error

    replacements = {
        spill_marker: "#",
        f"'{implicit_qualifier_marker}": "@'",
    } | {marker: escape for escape, marker in structured_markers.items()}
    tokens = _normalize_intersection_groups(
        _restore_source_spaces(shielded[1:], raw_tokens, replacements)
    )
    _validate_structure(tokens)
    return tokens


def _supports_exact_fast_path(formula: str) -> bool:
    """Identify formulas that need no lexical compatibility shielding."""
    return (
        "#" not in formula
        and "@'" not in formula
        and not any(character.isspace() for character in formula)
        and not any(escape in formula for escape in _STRUCTURED_ESCAPES)
    )


def _unused_marker(formula: str, label: str) -> str:
    number = 0
    while True:
        marker = f"__XLSP_{label}_{number}__"
        if marker not in formula:
            return marker
        number += 1


def _shield_modern_syntax(
    formula: str,
    spill_marker: str,
    implicit_qualifier_marker: str,
    structured_markers: dict[str, str],
) -> str:
    """Shield syntax that openpyxl 3.1.5 does not tokenize correctly."""
    result: list[str] = []
    single_quoted = False
    double_quoted = False
    bracket_depth = 0
    index = 0
    while index < len(formula):
        character = formula[index]
        if single_quoted:
            result.append(character)
            if character == "'":
                if index + 1 < len(formula) and formula[index + 1] == "'":
                    result.append("'")
                    index += 2
                    continue
                single_quoted = False
            index += 1
            continue
        if double_quoted:
            result.append(character)
            if character == '"':
                if index + 1 < len(formula) and formula[index + 1] == '"':
                    result.append('"')
                    index += 2
                    continue
                double_quoted = False
            index += 1
            continue
        if (
            bracket_depth
            and character == "'"
            and index + 1 < len(formula)
            and formula[index : index + 2] in structured_markers
        ):
            result.append(structured_markers[formula[index : index + 2]])
            index += 2
            continue
        if (
            character == "@"
            and bracket_depth == 0
            and index + 1 < len(formula)
            and formula[index + 1] == "'"
            and _quoted_qualifier_has_bang(formula, index + 1)
        ):
            # openpyxl rejects ``@'Sheet Name'!A1``.  Move a collision-free
            # marker just inside the quote for tokenization; restoration moves
            # it back in front without changing the resulting RANGE token.
            result.extend(("'", implicit_qualifier_marker))
            single_quoted = True
            index += 2
            continue
        if character == "'" and bracket_depth == 0:
            single_quoted = True
        elif character == '"' and bracket_depth == 0:
            double_quoted = True
        elif character == "[":
            bracket_depth += 1
        elif character == "]" and bracket_depth:
            bracket_depth -= 1
        elif character == "#" and bracket_depth == 0 and _is_postfix_spill(formula, index):
            result.append(spill_marker)
            index += 1
            continue
        result.append(character)
        index += 1
    return "".join(result)


def _quoted_qualifier_has_bang(formula: str, opening_quote: int) -> bool:
    index = opening_quote + 1
    while index < len(formula):
        if formula[index] != "'":
            index += 1
            continue
        if index + 1 < len(formula) and formula[index + 1] == "'":
            index += 2
            continue
        return index + 1 < len(formula) and formula[index + 1] == "!"
    return False


def _is_postfix_spill(formula: str, index: int) -> bool:
    if index == 0:
        return False
    previous = formula[index - 1]
    if not (previous.isalnum() or previous in "_.]"):
        return False
    if index + 1 == len(formula):
        return True
    following = formula[index + 1]
    return following.isspace() or following in "+-*/^&%=<>,;:)"


def _normalize_tokenizer_whitespace(formula: str) -> str:
    """Prevent openpyxl from reordering newline whitespace before operands."""
    result: list[str] = []
    single_quoted = False
    double_quoted = False
    bracket_depth = 0
    index = 0
    while index < len(formula):
        character = formula[index]
        if single_quoted:
            result.append(character)
            if character == "'":
                if index + 1 < len(formula) and formula[index + 1] == "'":
                    result.append("'")
                    index += 2
                    continue
                single_quoted = False
        elif double_quoted:
            result.append(character)
            if character == '"':
                if index + 1 < len(formula) and formula[index + 1] == '"':
                    result.append('"')
                    index += 2
                    continue
                double_quoted = False
        elif character == "'" and bracket_depth == 0:
            single_quoted = True
            result.append(character)
        elif character == '"' and bracket_depth == 0:
            double_quoted = True
            result.append(character)
        elif character == "[":
            bracket_depth += 1
            result.append(character)
        elif character == "]" and bracket_depth:
            bracket_depth -= 1
            result.append(character)
        elif character in {"\n", "\r"} and bracket_depth == 0:
            result.append(" ")
        else:
            result.append(character)
        index += 1
    return "".join(result)


def _restore_source_spaces(
    body: str,
    raw_tokens: list[Token],
    replacements: dict[str, str],
) -> tuple[FormulaToken, ...]:
    """Restore space runs collapsed by openpyxl while checking lexical fidelity."""
    position = 0
    restored: list[FormulaToken] = []
    for raw in raw_tokens:
        value = raw.value
        token_type = raw.type
        subtype = raw.subtype
        if token_type == WHITESPACE_TOKEN_TYPE:
            start = position
            while position < len(body) and body[position] in {" ", "\n", "\r"}:
                position += 1
            if position == start:
                raise FormulaSyntaxError("tokenizer whitespace did not match formula source")
            value = body[start:position]
        else:
            if not body.startswith(value, position):
                raise FormulaSyntaxError("tokenizer output did not round-trip to formula source")
            position += len(value)
        for marker, original in replacements.items():
            value = value.replace(marker, original)
        restored.append(FormulaToken(value, token_type, subtype))
    if position != len(body):
        raise FormulaSyntaxError("tokenizer did not consume the complete formula")
    return tuple(restored)


def _normalize_intersection_groups(
    tokens: tuple[FormulaToken, ...],
) -> tuple[FormulaToken, ...]:
    """Represent ``@(...)`` as a prefix operator plus ordinary grouping.

    openpyxl reports ``@(`` as an empty function call and may absorb a left
    range endpoint into the same token (``A1:@(``).  Splitting that spelling
    keeps source fidelity while preventing an empty callable and lets every
    downstream pass see and translate both range endpoints.
    """
    normalized: list[FormulaToken] = []
    stack: list[tuple[str, bool]] = []
    for token in tokens:
        if token.type in _OPEN_TYPES and token.subtype == "OPEN":
            intersection = token.type == "FUNC" and _append_intersection_open(
                token.value,
                normalized,
            )
            stack.append((token.type, intersection))
            if not intersection:
                normalized.append(token)
            continue
        if token.type in _OPEN_TYPES and token.subtype == "CLOSE":
            if stack and stack[-1][0] == token.type:
                _opening_type, intersection = stack.pop()
                if intersection:
                    normalized.append(FormulaToken(token.value, "PAREN", "CLOSE"))
                    continue
            normalized.append(token)
            continue
        normalized.append(token)
    return tuple(normalized)


def _append_intersection_open(value: str, output: list[FormulaToken]) -> bool:
    if not value.endswith("("):
        return False
    body = value[:-1]
    if body == "@":
        prefix = ""
        has_range_operator = False
    elif body.endswith(":@"):
        prefix = body[:-2]
        has_range_operator = True
    else:
        return False
    if prefix:
        output.append(FormulaToken(prefix, "OPERAND", "RANGE"))
    if has_range_operator:
        output.append(FormulaToken(":", "OPERAND", "RANGE"))
    output.append(FormulaToken("@", "OPERATOR-PREFIX", ""))
    output.append(FormulaToken("(", "PAREN", "OPEN"))
    return True


def _validate_structure(tokens: tuple[FormulaToken, ...]) -> None:
    stack: list[str] = []
    significant = False
    for token in tokens:
        if token.type != WHITESPACE_TOKEN_TYPE:
            significant = True
        if token.type in _OPEN_TYPES and token.subtype == "OPEN":
            stack.append(token.type)
        elif token.type in _OPEN_TYPES and token.subtype == "CLOSE":
            if not stack or stack[-1] != token.type:
                raise FormulaSyntaxError("formula contains mismatched closing tokens")
            stack.pop()
    if stack:
        raise FormulaSyntaxError("formula contains unclosed structural tokens")
    if not significant:
        raise FormulaSyntaxError("formula is empty")


__all__ = [
    "WHITESPACE_TOKEN_TYPE",
    "FormulaSyntaxError",
    "FormulaToken",
    "tokenize_formula",
]
