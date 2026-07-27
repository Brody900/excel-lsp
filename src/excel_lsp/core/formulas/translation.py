"""Modern-formula-safe A1 translation for shared formula groups."""

from __future__ import annotations

from typing import cast

from openpyxl.formula.translate import Translator, TranslatorError

from excel_lsp.core.formulas.a1 import (
    CellRef,
    ParsedRangeEndpoint,
    ParsedReferenceRange,
    parse_a1_reference,
    parse_modern_a1_range,
    parse_reference_range,
)
from excel_lsp.core.formulas.tokens import FormulaToken, tokenize_formula


def translate_a1_formula(
    formula: str,
    *,
    origin: CellRef,
    target: CellRef,
    preserve_coordinate_spills: bool = False,
) -> str:
    """Translate relative A1 references while preserving modern formula syntax."""
    tokens = tokenize_formula(formula)
    return "=" + "".join(
        _translate_token(
            token,
            origin,
            target,
            preserve_coordinate_spills=preserve_coordinate_spills,
        )
        for token in tokens
    )


def _translate_token(
    token: FormulaToken,
    origin: CellRef,
    target: CellRef,
    *,
    preserve_coordinate_spills: bool,
) -> str:
    if token.type == "FUNC" and token.subtype == "OPEN":
        body = token.value[:-1]
        if ":" not in body:
            return token.value
        reference, function_name = body.rsplit(":", 1)
        translated = _translate_reference_text(
            reference,
            origin,
            target,
            preserve_coordinate_spills=preserve_coordinate_spills,
        )
        return f"{translated}:{function_name}("
    if token.type == "OPERAND" and token.subtype == "RANGE":
        leading_colon = token.value.startswith(":")
        reference = token.value[1:] if leading_colon else token.value
        translated = _translate_reference_text(
            reference,
            origin,
            target,
            preserve_coordinate_spills=preserve_coordinate_spills,
        )
        return f":{translated}" if leading_colon else translated
    return token.value


def _translate_reference_text(
    value: str,
    origin: CellRef,
    target: CellRef,
    *,
    preserve_coordinate_spills: bool,
) -> str:
    modern_range = parse_modern_a1_range(value, origin)
    if modern_range is not None:
        left = _translate_modern_endpoint(
            modern_range.left_original,
            implicit=modern_range.left_implicit,
            spill=modern_range.left_spill,
            origin=origin,
            target=target,
            preserve_coordinate_spills=preserve_coordinate_spills,
        )
        right = _translate_modern_endpoint(
            modern_range.right_original,
            implicit=modern_range.right_implicit,
            spill=modern_range.right_spill,
            origin=origin,
            target=target,
            preserve_coordinate_spills=preserve_coordinate_spills,
        )
        return f"{left}:{right}"
    reference_range = parse_reference_range(value, origin)
    if reference_range is not None:
        return _translate_reference_range(
            reference_range,
            origin,
            target,
            preserve_coordinate_spills=preserve_coordinate_spills,
        )
    implicit_intersection = value.startswith("@")
    candidate = value[1:] if implicit_intersection else value
    spill = candidate.endswith("#")
    if spill:
        candidate = candidate[:-1]
    parsed = parse_a1_reference(candidate, origin)
    if parsed is None:
        return value
    if spill and preserve_coordinate_spills:
        return value
    translated = _translate_a1_text(candidate, origin, target)
    return f"{'@' if implicit_intersection else ''}{translated}{'#' if spill else ''}"


def _translate_reference_range(
    parsed: ParsedReferenceRange,
    origin: CellRef,
    target: CellRef,
    *,
    preserve_coordinate_spills: bool,
) -> str:
    left = _translate_mixed_endpoint(
        parsed.left,
        origin,
        target,
        preserve_coordinate_spills=preserve_coordinate_spills,
    )
    right = _translate_mixed_endpoint(
        parsed.right,
        origin,
        target,
        preserve_coordinate_spills=preserve_coordinate_spills,
    )
    return f"{left}:{right}"


def _translate_mixed_endpoint(
    endpoint: ParsedRangeEndpoint,
    origin: CellRef,
    target: CellRef,
    *,
    preserve_coordinate_spills: bool,
) -> str:
    parsed = endpoint.parsed_a1
    if parsed is None:
        return endpoint.original
    if endpoint.spill and preserve_coordinate_spills:
        return endpoint.original
    translated = _translate_a1_text(endpoint.core, origin, target)
    return f"{'@' if endpoint.implicit else ''}{translated}{'#' if endpoint.spill else ''}"


def _translate_modern_endpoint(
    original: str,
    *,
    implicit: bool,
    spill: bool,
    origin: CellRef,
    target: CellRef,
    preserve_coordinate_spills: bool,
) -> str:
    if spill and preserve_coordinate_spills:
        return original
    core = original[1:] if implicit else original
    if spill:
        core = core[:-1]
    translated = _translate_a1_text(core, origin, target)
    return f"{'@' if implicit else ''}{translated}{'#' if spill else ''}"


def _translate_a1_text(text: str, origin: CellRef, target: CellRef) -> str:
    """Translate one already-validated A1 area without changing absolute case."""
    try:
        translated = cast(
            str,
            Translator.translate_range(  # pyright: ignore[reportUnknownMemberType]
                text,
                target.row - origin.row,
                target.col - origin.col,
            ),
        )
    except (TranslatorError, ValueError) as error:
        raise ValueError("translated reference exceeds Excel bounds") from error
    if parse_a1_reference(translated, target) is None:
        raise ValueError("translated reference exceeds Excel bounds")
    return translated


__all__ = ["translate_a1_formula"]
