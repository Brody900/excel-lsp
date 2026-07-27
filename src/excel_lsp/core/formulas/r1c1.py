"""Canonical A1-to-R1C1 formula normalization."""

from __future__ import annotations

from excel_lsp.core.formulas.a1 import (
    CellRef,
    ParsedRangeEndpoint,
    ParsedReferenceRange,
    decorate_modern_a1_range,
    parse_a1_reference,
    parse_modern_a1_range,
    parse_reference_range,
    render_r1c1_reference,
)
from excel_lsp.core.formulas.tokens import FormulaToken, tokenize_formula


def to_r1c1(formula: str, anchor: CellRef) -> str:
    """Rewrite A1 operands relative to ``anchor`` and preserve all other text."""
    tokens = tokenize_formula(formula)
    return "=" + "".join(_normalize_token(token, anchor) for token in tokens)


def _normalize_token(token: FormulaToken, anchor: CellRef) -> str:
    if token.type == "FUNC" and token.subtype == "OPEN":
        body = token.value[:-1]
        if ":" not in body:
            return token.value
        reference, function_name = body.rsplit(":", 1)
        normalized = _normalize_reference_text(reference, anchor)
        return f"{normalized}:{function_name}("
    if token.type == "OPERAND" and token.subtype == "RANGE":
        leading_colon = token.value.startswith(":")
        reference = token.value[1:] if leading_colon else token.value
        normalized = _normalize_reference_text(reference, anchor)
        return f":{normalized}" if leading_colon else normalized
    return token.value


def _normalize_reference_text(value: str, anchor: CellRef) -> str:
    modern_range = parse_modern_a1_range(value, anchor)
    if modern_range is not None:
        rendered = render_r1c1_reference(
            modern_range.parsed,
            preserve_range=True,
        )
        return decorate_modern_a1_range(
            rendered,
            modern_range,
            preserve_spill_endpoints=True,
        )
    reference_range = parse_reference_range(value, anchor)
    if reference_range is not None:
        return _normalize_reference_range(reference_range)
    # Section 5.4 deliberately keeps spill references verbatim.  Their anchor
    # is resolved separately by the reference-extraction layer.
    if value.endswith("#"):
        return value
    implicit_intersection = value.startswith("@")
    candidate = value[1:] if implicit_intersection else value
    parsed = parse_a1_reference(candidate, anchor)
    if parsed is None:
        return value
    local = candidate[len(parsed.qualifier) :]
    rendered = render_r1c1_reference(parsed, preserve_range=":" in local)
    return f"@{rendered}" if implicit_intersection else rendered


def _normalize_reference_range(parsed: ParsedReferenceRange) -> str:
    return f"{_normalize_mixed_endpoint(parsed.left)}:{_normalize_mixed_endpoint(parsed.right)}"


def _normalize_mixed_endpoint(endpoint: ParsedRangeEndpoint) -> str:
    parsed = endpoint.parsed_a1
    # Section 5.4 keeps spill anchors verbatim even when the other endpoint is
    # a name.  The shared-formula translator handles their physical shift.
    if parsed is None or endpoint.spill:
        return endpoint.original
    local = endpoint.core[len(parsed.qualifier) :]
    rendered = render_r1c1_reference(parsed, preserve_range=":" in local)
    return f"{'@' if endpoint.implicit else ''}{rendered}"


__all__ = ["to_r1c1"]
