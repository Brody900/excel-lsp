"""Modern-formula-safe A1 translation for shared formula groups."""

from __future__ import annotations

from dataclasses import dataclass
from typing import cast

from openpyxl.formula.translate import Translator, TranslatorError

from excel_lsp.core.formulas.a1 import (
    CellRef,
    ParsedA1Reference,
    ParsedModernA1Range,
    ParsedRangeEndpoint,
    ParsedReferenceRange,
    decorate_modern_a1_range,
    parse_a1_reference,
    parse_modern_a1_range,
    parse_reference_range,
    render_a1_reference,
)
from excel_lsp.core.formulas.tokens import FormulaToken, tokenize_formula


@dataclass(frozen=True, slots=True)
class A1TranslationPlan:
    """One tokenized A1 formula reusable across a shared-formula group."""

    formula: str
    origin: CellRef
    tokens: tuple[_PlannedToken, ...]

    @classmethod
    def compile(cls, formula: str, *, origin: CellRef) -> A1TranslationPlan:
        return cls(
            formula=formula,
            origin=origin,
            tokens=tuple(
                _PlannedToken.compile(token, origin) for token in tokenize_formula(formula)
            ),
        )

    def translate(
        self,
        *,
        target: CellRef,
        preserve_coordinate_spills: bool = False,
    ) -> str:
        return "=" + "".join(
            token.translate(
                origin=self.origin,
                target=target,
                preserve_coordinate_spills=preserve_coordinate_spills,
            )
            for token in self.tokens
        )


@dataclass(frozen=True, slots=True)
class _PlannedToken:
    original: FormulaToken
    reference: _PlannedReference | None = None
    prefix: str = ""
    suffix: str = ""

    @classmethod
    def compile(cls, token: FormulaToken, origin: CellRef) -> _PlannedToken:
        if token.type == "FUNC" and token.subtype == "OPEN":
            body = token.value[:-1]
            if ":" in body:
                reference, function_name = body.rsplit(":", 1)
                return cls(
                    token,
                    _PlannedReference.compile(reference, origin),
                    suffix=f":{function_name}(",
                )
        if token.type == "OPERAND" and token.subtype == "RANGE":
            leading_colon = token.value.startswith(":")
            reference = token.value[1:] if leading_colon else token.value
            return cls(
                token,
                _PlannedReference.compile(reference, origin),
                prefix=":" if leading_colon else "",
            )
        return cls(token)

    def translate(
        self,
        *,
        origin: CellRef,
        target: CellRef,
        preserve_coordinate_spills: bool,
    ) -> str:
        if self.reference is None:
            return self.original.value
        return (
            self.prefix
            + self.reference.translate(
                origin=origin,
                target=target,
                preserve_coordinate_spills=preserve_coordinate_spills,
            )
            + self.suffix
        )


@dataclass(frozen=True, slots=True)
class _PlannedReference:
    original: str
    parsed: ParsedA1Reference | None = None
    modern: ParsedModernA1Range | None = None
    mixed: ParsedReferenceRange | None = None
    implicit: bool = False
    spill: bool = False
    fast: bool = False

    @classmethod
    def compile(cls, value: str, origin: CellRef) -> _PlannedReference:
        # Rendering parsed geometry canonicalizes A1 columns to uppercase.  The
        # compatibility path intentionally retains lowercase-directional
        # behavior, names, and other source spellings exactly.
        fast = not any(character.isascii() and character.islower() for character in value)
        if not fast:
            return cls(value)
        modern = parse_modern_a1_range(value, origin)
        if modern is not None:
            return cls(value, modern=modern, fast=True)
        mixed = parse_reference_range(value, origin)
        if mixed is not None:
            return cls(value, mixed=mixed, fast=True)
        implicit = value.startswith("@")
        candidate = value[1:] if implicit else value
        spill = candidate.endswith("#")
        if spill:
            candidate = candidate[:-1]
        parsed = parse_a1_reference(candidate, origin)
        if parsed is None:
            return cls(value)
        return cls(value, parsed=parsed, implicit=implicit, spill=spill, fast=True)

    def translate(
        self,
        *,
        origin: CellRef,
        target: CellRef,
        preserve_coordinate_spills: bool,
    ) -> str:
        if not self.fast:
            return _translate_reference_text(
                self.original,
                origin,
                target,
                preserve_coordinate_spills=preserve_coordinate_spills,
            )
        if self.modern is not None:
            rendered = render_a1_reference(self.modern.parsed, target, preserve_range=True)
            return decorate_modern_a1_range(
                rendered,
                self.modern,
                preserve_spill_endpoints=preserve_coordinate_spills,
            )
        if self.mixed is not None:
            return _render_mixed_reference(
                self.mixed,
                target,
                preserve_coordinate_spills=preserve_coordinate_spills,
            )
        if self.parsed is None:
            return self.original
        if self.spill and preserve_coordinate_spills:
            return self.original
        rendered = render_a1_reference(self.parsed, target)
        return f"{'@' if self.implicit else ''}{rendered}{'#' if self.spill else ''}"


def _render_mixed_reference(
    mixed: ParsedReferenceRange,
    target: CellRef,
    *,
    preserve_coordinate_spills: bool,
) -> str:
    return ":".join(
        _render_mixed_endpoint(
            endpoint,
            target,
            preserve_coordinate_spills=preserve_coordinate_spills,
        )
        for endpoint in (mixed.left, mixed.right)
    )


def _render_mixed_endpoint(
    endpoint: ParsedRangeEndpoint,
    target: CellRef,
    *,
    preserve_coordinate_spills: bool,
) -> str:
    if endpoint.parsed_a1 is None:
        return endpoint.original
    if endpoint.spill and preserve_coordinate_spills:
        return endpoint.original
    rendered = render_a1_reference(endpoint.parsed_a1, target)
    return f"{'@' if endpoint.implicit else ''}{rendered}{'#' if endpoint.spill else ''}"


def translate_a1_formula(
    formula: str,
    *,
    origin: CellRef,
    target: CellRef,
    preserve_coordinate_spills: bool = False,
) -> str:
    """Translate relative A1 references while preserving modern formula syntax."""
    return A1TranslationPlan.compile(formula, origin=origin).translate(
        target=target,
        preserve_coordinate_spills=preserve_coordinate_spills,
    )


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


__all__ = ["A1TranslationPlan", "translate_a1_formula"]
