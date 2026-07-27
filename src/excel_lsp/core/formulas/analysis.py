"""Formula-level reference extraction, lexical scope, and function flags."""

from __future__ import annotations

from dataclasses import dataclass, field

from excel_lsp.core.formulas.a1 import (
    AxisTerm,
    CellRef,
    ParsedA1Reference,
    ReferenceGeometry,
    parse_a1_reference,
    parse_reference_range_candidates,
    render_a1_reference,
)
from excel_lsp.core.formulas.functions import (
    ALWAYS_DYNAMIC_REFERENCE_FUNCTIONS,
    BUILTIN_FUNCTIONS,
    VOLATILE_FUNCTIONS,
    compatibility_function_identifier,
    normalize_function_name,
)
from excel_lsp.core.formulas.references import (
    ExtractedReference,
    FormulaAnchor,
    FormulaIssue,
    NameKey,
    ReferenceClassification,
    ReferenceContext,
    classify_ref,
    name_expansion_scope,
)
from excel_lsp.core.formulas.tokens import (
    WHITESPACE_TOKEN_TYPE,
    FormulaSyntaxError,
    FormulaToken,
    tokenize_formula,
)
from excel_lsp.core.models import DefinedName

_LOCAL_PARAMETER_PREFIX = "_xlpm."


def _new_bindings() -> set[str]:
    return set()


def _new_references() -> list[ExtractedReference]:
    return []


def _new_issues() -> list[FormulaIssue]:
    return []


def _new_function_calls() -> list[str]:
    return []


@dataclass(frozen=True, slots=True)
class ReferenceUse:
    """One direct formula operand and its anchor-time classification."""

    token: str
    function_position: bool
    classification: ReferenceClassification


@dataclass(frozen=True, slots=True)
class FormulaAnalysis:
    """Complete semantic extraction result for one stored formula."""

    references: tuple[ExtractedReference, ...]
    function_calls: tuple[str, ...]
    volatile: bool
    opaque: bool
    issues: tuple[FormulaIssue, ...]
    reference_uses: tuple[ReferenceUse, ...]
    returns_reference: bool = False
    dynamic_result_labels: tuple[str, ...] = ()


def _new_reference_uses() -> list[ReferenceUse]:
    return []


def _new_index_set() -> set[int]:
    return set()


def _new_lambda_call_set() -> set[tuple[int, int]]:
    return set()


def _new_dynamic_callable_groups() -> dict[int, str]:
    return {}


@dataclass(slots=True)
class _Frame:
    token_type: str
    name: str | None = None
    intrinsic_name: str | None = None
    source_token: str = ""
    arg_index: int = 0
    pending_declaration: FormulaToken | None = None
    pending_let_binding: str | None = None
    arg_has_other: bool = False
    bindings: set[str] = field(default_factory=_new_bindings)
    declaration_names: set[str] = field(default_factory=_new_bindings)
    pending_let_declaration: str | None = None
    lexical_composite: bool = False
    choose_branch_reference: bool = False
    choose_current_reference: bool = False
    choose_current_other: bool = False
    choose_union_pending: bool = False
    choose_union_invalid: bool = False
    dynamic_emitted: bool = False
    returns_reference: bool = False

    @property
    def special(self) -> bool:
        return self.intrinsic_name in {"LET", "LAMBDA"}


@dataclass(slots=True)
class _Accumulator:
    references: list[ExtractedReference] = field(default_factory=_new_references)
    issues: list[FormulaIssue] = field(default_factory=_new_issues)
    function_calls: list[str] = field(default_factory=_new_function_calls)
    reference_uses: list[ReferenceUse] = field(default_factory=_new_reference_uses)
    volatile: bool = False
    opaque: bool = False

    def accept(
        self,
        classification: ReferenceClassification,
        *,
        source_token: str | None = None,
        function_position: bool = False,
    ) -> None:
        self.references.extend(classification.references)
        self.issues.extend(classification.issues)
        self.function_calls.extend(classification.function_calls)
        self.volatile = self.volatile or classification.volatile
        self.opaque = self.opaque or classification.opaque
        if source_token is not None:
            self.reference_uses.append(
                ReferenceUse(source_token, function_position, classification)
            )


@dataclass(slots=True)
class _ReferenceInference:
    """Reference-result facts discovered with lexical binding types."""

    reference_functions: set[int] = field(default_factory=_new_index_set)
    dynamic_choices: set[int] = field(default_factory=_new_index_set)
    active_lambda_calls: set[tuple[int, int]] = field(default_factory=_new_lambda_call_set)
    dynamic_callable_groups: dict[int, str] = field(default_factory=_new_dynamic_callable_groups)


@dataclass(slots=True)
class _CallableBinding:
    """Marker for a statically retained callable value."""


@dataclass(slots=True)
class _LambdaBinding(_CallableBinding):
    """One lexical LAMBDA value together with its captured binding types."""

    tokens: tuple[FormulaToken, ...]
    matching: dict[int, int]
    opening: int
    parameter_bounds: tuple[tuple[int, int] | None, ...]
    body: tuple[int, int] | None
    closure: dict[str, bool | _CallableBinding]
    name_stack: tuple[NameKey, ...]
    inference: _ReferenceInference


@dataclass(slots=True)
class _CallableChoice(_CallableBinding):
    """Callable alternatives returned by a selector such as CHOOSE or IF."""

    alternatives: tuple[_CallableBinding, ...]
    truncated: bool = False


_MAX_CALLABLE_CHOICE_ALTERNATIVES = 32


def analyze_formula(
    formula: str,
    *,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...] = (),
) -> FormulaAnalysis:
    """Extract all static references while containing per-formula failures."""
    with name_expansion_scope():
        try:
            tokens = tokenize_formula(formula)
            return _analyze_tokens(tokens, anchor=anchor, context=context, name_stack=name_stack)
        except (FormulaSyntaxError, ValueError, TypeError) as error:
            return _failed_analysis(formula, str(error))
        except Exception as error:  # pragma: no cover - invariant I7 containment backstop
            return _failed_analysis(formula, f"unexpected formula parser failure: {error}")


def _analyze_tokens(
    tokens: tuple[FormulaToken, ...],
    *,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> FormulaAnalysis:
    accumulator = _Accumulator()
    frames: list[_Frame] = []
    tokens = _fold_static_range_operators(
        tokens,
        anchor,
        context,
        name_stack,
    )
    matching_groups = _matching_groups(tokens)
    inferred, returns_reference = _infer_reference_types(
        tokens,
        matching_groups,
        anchor,
        context,
        name_stack,
    )
    reference_functions = frozenset(inferred.reference_functions)
    dynamic_indexes = _dynamic_index_positions(
        tokens,
        matching_groups,
        frozenset(
            index
            for index in reference_functions
            if normalize_function_name(tokens[index].value) == "INDEX"
        ),
        reference_functions,
    )
    dynamic_reference_endpoints = set(
        _reference_function_range_endpoint_positions(
            tokens,
            matching_groups,
            reference_functions,
        )
    )
    intrinsic_dynamic_functions = (
        dynamic_indexes
        | frozenset(inferred.dynamic_choices)
        | frozenset(
            index
            for index in reference_functions
            if normalize_function_name(tokens[index].value) in ALWAYS_DYNAMIC_REFERENCE_FUNCTIONS
        )
    )
    for endpoint in tuple(dynamic_reference_endpoints):
        if normalize_function_name(tokens[endpoint].value) not in {"LET", "LAMBDA", "SINGLE"}:
            continue
        closing = matching_groups.get(endpoint)
        if closing is not None and any(
            endpoint < nested < closing for nested in intrinsic_dynamic_functions
        ):
            dynamic_reference_endpoints.remove(endpoint)
    dynamic_reference_functions = dynamic_indexes | frozenset(dynamic_reference_endpoints)
    for group_index, label in sorted(inferred.dynamic_callable_groups.items()):
        _emit_dynamic_label(
            accumulator,
            source_token=tokens[group_index].value,
            label=label,
        )

    lexical_intersections: set[int] = set()
    for token_index, token in enumerate(tokens):
        if token.type == WHITESPACE_TOKEN_TYPE:
            marker = _lexical_intersection_marker(
                tokens,
                matching_groups,
                token_index,
                frames,
            )
            if marker is not None:
                lexical_intersections.add(token_index)
                accumulator.references.append(
                    ExtractedReference(marker, None, None, None, "opaque:ref")
                )
                accumulator.opaque = True
            continue
        if token.subtype == "CLOSE" and token.type in {"FUNC", "PAREN", "ARRAY"}:
            frame = frames[-1]
            if frame.token_type == "FUNC":
                _finish_special_function(
                    frame,
                    accumulator,
                    frames,
                    anchor,
                    context,
                    name_stack,
                )
                _finish_choose_argument(frame)
                if frame.intrinsic_name == "CHOOSE" and frame.choose_branch_reference:
                    _emit_dynamic(accumulator, frame)
                frames.pop()
                parent = _active_function_frame(frames)
                if (
                    parent is not None
                    and parent.intrinsic_name == "CHOOSE"
                    and parent.arg_index >= 1
                    and _frame_returns_reference(frame)
                ):
                    parent.choose_current_reference = True
                    parent.choose_union_pending = False
            else:
                frames.pop()
            continue

        if token.type == "SEP":
            if frames and frames[-1].token_type == "FUNC" and token.subtype == "ARG":
                frame = frames[-1]
                _finish_argument(
                    frame,
                    accumulator,
                    frames,
                    anchor,
                    context,
                    name_stack,
                )
                _finish_choose_argument(frame)
                frame.arg_index += 1
                frame.pending_declaration = None
                frame.arg_has_other = False
            continue

        active_function = _active_function_frame(frames)
        if (
            active_function is not None
            and active_function.intrinsic_name == "CHOOSE"
            and active_function.arg_index >= 1
            and token.type == "OPERATOR-INFIX"
            and token.value == ","
        ):
            if (
                not active_function.choose_current_reference
                or active_function.choose_current_other
                or active_function.choose_union_pending
            ):
                active_function.choose_union_invalid = True
            active_function.choose_union_pending = True
        if (
            active_function is not None
            and active_function.intrinsic_name == "CHOOSE"
            and active_function.arg_index >= 1
            and not _token_can_start_reference_expression(
                token,
                function_returns_reference=(
                    token_index in reference_functions
                    or (
                        token.type == "FUNC"
                        and token.subtype == "OPEN"
                        and normalize_function_name(token.value)
                        in ALWAYS_DYNAMIC_REFERENCE_FUNCTIONS
                    )
                ),
            )
        ):
            active_function.choose_current_other = True

        special = frames[-1] if frames and frames[-1].token_type == "FUNC" else None
        if (
            token.type == "OPERAND"
            and token.subtype == "RANGE"
            and special is not None
            and special.special
            and _argument_can_declare(special)
            and not special.arg_has_other
            and special.pending_declaration is None
            and _is_declaration_identifier(token.value)
        ):
            special.pending_declaration = token
            continue

        if special is not None and special.special:
            _flush_pending_as_expression(
                special,
                accumulator,
                frames,
                anchor,
                context,
                name_stack,
            )
            special.arg_has_other = True

        if token.subtype == "OPEN" and token.type in {"FUNC", "PAREN", "ARRAY"}:
            if token.type != "FUNC":
                frames.append(
                    _Frame(
                        token.type,
                        lexical_composite=(
                            any(frame.lexical_composite for frame in frames)
                            or _group_is_lexical_colon_endpoint(
                                tokens,
                                matching_groups,
                                token_index,
                                frames,
                            )
                        ),
                    )
                )
                continue

            prefix = _function_reference_prefix(token.value)
            marker: str | None = None
            if _is_synthetic_range_group_open(token):
                marker = _lexical_synthetic_colon_marker(
                    tokens,
                    matching_groups,
                    token_index,
                    prefix,
                    frames,
                )
                if marker is not None:
                    accumulator.references.append(
                        ExtractedReference(marker, None, None, None, "opaque:ref")
                    )
                    accumulator.opaque = True
            if prefix:
                _process_range(
                    FormulaToken(prefix, "OPERAND", "RANGE"),
                    accumulator,
                    frames,
                    anchor,
                    context,
                    name_stack,
                    lexical_endpoint=marker is not None,
                )
            callable_identifier = compatibility_function_identifier(token.value)
            if not callable_identifier and _is_synthetic_range_group_open(token):
                frames.append(_Frame("PAREN", lexical_composite=marker is not None))
                continue

            name = normalize_function_name(token.value)
            bound_callable = name not in BUILTIN_FUNCTIONS and _is_bound(
                callable_identifier, frames
            )
            display_name = _callable_label(token.value) if bound_callable else name
            frame = _Frame(
                "FUNC",
                name=display_name,
                intrinsic_name=name,
                source_token=token.value,
                returns_reference=(
                    name in ALWAYS_DYNAMIC_REFERENCE_FUNCTIONS or token_index in reference_functions
                ),
            )
            accumulator.function_calls.append(display_name)
            if name in VOLATILE_FUNCTIONS:
                accumulator.volatile = True
            if name not in BUILTIN_FUNCTIONS and not bound_callable:
                classification = classify_ref(
                    callable_identifier,
                    anchor=anchor,
                    context=context,
                    function_position=True,
                    name_stack=name_stack,
                )
                accumulator.accept(
                    classification,
                    source_token=callable_identifier,
                    function_position=True,
                )
            if name in ALWAYS_DYNAMIC_REFERENCE_FUNCTIONS:
                _emit_dynamic(accumulator, frame)
            if token_index in inferred.dynamic_choices:
                _emit_dynamic(accumulator, frame)
            if token_index in dynamic_reference_functions:
                _emit_dynamic(accumulator, frame)
            frames.append(frame)
            continue

        if token.type == "OPERAND" and token.subtype == "RANGE":
            if token.value == ":":
                continue
            marker = _lexical_suffix_colon_marker(
                tokens,
                matching_groups,
                token_index,
                frames,
            )
            if marker is not None:
                accumulator.references.append(
                    ExtractedReference(marker, None, None, None, "opaque:ref")
                )
                accumulator.opaque = True
            _process_range(
                token,
                accumulator,
                frames,
                anchor,
                context,
                name_stack,
                lexical_endpoint=(
                    marker is not None or any(frame.lexical_composite for frame in frames)
                ),
            )
            continue

    intersection = _unfolded_intersection_metadata(
        tokens,
        anchor,
        context,
        name_stack,
        skip_indexes=frozenset(lexical_intersections),
    )
    if intersection is not None:
        intersection_accumulator = _Accumulator()
        dynamic_events, needs_generic_marker = intersection
        for source_token, label in dynamic_events:
            _emit_dynamic_label(
                intersection_accumulator,
                source_token=source_token,
                label=label,
            )
        if needs_generic_marker:
            intersection_accumulator.references.append(
                ExtractedReference(" ", None, None, None, "opaque:ref")
            )
            intersection_accumulator.opaque = True
        accumulator.references[:0] = intersection_accumulator.references
        accumulator.issues[:0] = intersection_accumulator.issues
        accumulator.opaque = accumulator.opaque or intersection_accumulator.opaque

    dynamic_result_labels = _result_dynamic_labels(
        tokens,
        matching_groups,
        inferred,
        accumulator.reference_uses,
        returns_reference,
    )
    return FormulaAnalysis(
        tuple(accumulator.references),
        tuple(accumulator.function_calls),
        accumulator.volatile,
        accumulator.opaque,
        tuple(accumulator.issues),
        tuple(accumulator.reference_uses),
        returns_reference,
        dynamic_result_labels,
    )


def _finish_argument(
    frame: _Frame,
    accumulator: _Accumulator,
    frames: list[_Frame],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> None:
    if not frame.special:
        return
    if _argument_can_declare(frame):
        pending = frame.pending_declaration
        if pending is not None and not frame.arg_has_other:
            binding = _local_binding_key(pending.value)
            declaration = _local_display_key(pending.value)
            if declaration in frame.declaration_names:
                _mark_parse_problem(
                    accumulator,
                    pending.value,
                    f"{frame.name} declaration names must be unique.",
                )
            if frame.intrinsic_name == "LET":
                frame.pending_let_binding = binding
                frame.pending_let_declaration = declaration
            else:
                frame.bindings.add(binding)
                frame.declaration_names.add(declaration)
            return
        _mark_parse_problem(
            accumulator,
            frame.source_token,
            f"{frame.name} declaration argument is not one identifier.",
        )
        return
    _flush_pending_as_expression(
        frame,
        accumulator,
        frames,
        anchor,
        context,
        name_stack,
    )
    if frame.intrinsic_name == "LET" and frame.arg_index % 2 == 1:
        binding = frame.pending_let_binding
        if binding is not None:
            frame.bindings.add(binding)
            declaration = frame.pending_let_declaration
            if declaration is not None:
                frame.declaration_names.add(declaration)
            frame.pending_let_binding = None
            frame.pending_let_declaration = None


def _finish_special_function(
    frame: _Frame,
    accumulator: _Accumulator,
    frames: list[_Frame],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> None:
    if not frame.special:
        return
    _flush_pending_as_expression(
        frame,
        accumulator,
        frames,
        anchor,
        context,
        name_stack,
    )
    empty_final_let = (
        frame.intrinsic_name == "LET" and frame.arg_index >= 2 and frame.arg_index % 2 == 0
    )
    if not empty_final_let and not frame.arg_has_other and frame.pending_declaration is None:
        _mark_parse_problem(
            accumulator,
            frame.source_token,
            f"{frame.name} has an empty final expression.",
        )
    if frame.intrinsic_name == "LET" and (frame.arg_index < 2 or frame.arg_index % 2 == 1):
        _mark_parse_problem(
            accumulator,
            frame.source_token,
            "LET requires name/value pairs followed by one calculation.",
        )


def _flush_pending_as_expression(
    frame: _Frame,
    accumulator: _Accumulator,
    frames: list[_Frame],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> None:
    pending = frame.pending_declaration
    if pending is None:
        return
    frame.pending_declaration = None
    _process_range(pending, accumulator, frames, anchor, context, name_stack)
    frame.arg_has_other = True


def _process_range(
    token: FormulaToken,
    accumulator: _Accumulator,
    frames: list[_Frame],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
    *,
    lexical_endpoint: bool = False,
) -> None:
    value = token.value[1:] if token.value.startswith(":") else token.value
    if not value:
        return
    composite = _lexical_composite_endpoints(value, frames, anchor)
    if composite is not None:
        for endpoint in composite:
            if _is_bound(endpoint, frames):
                continue
            accumulator.accept(
                _classify_lexical_endpoint(
                    endpoint,
                    anchor=anchor,
                    context=context,
                    name_stack=name_stack,
                ),
                source_token=endpoint,
            )
        accumulator.references.append(ExtractedReference(value, None, None, None, "opaque:ref"))
        accumulator.opaque = True
        return
    if _is_bound(value, frames):
        active_function = _active_function_frame(frames)
        if (
            active_function is not None
            and active_function.intrinsic_name == "CHOOSE"
            and active_function.arg_index >= 1
        ):
            active_function.choose_current_other = True
            active_function.choose_union_pending = False
        return
    classification = (
        _classify_lexical_endpoint(
            value,
            anchor=anchor,
            context=context,
            name_stack=name_stack,
        )
        if lexical_endpoint
        else classify_ref(
            value,
            anchor=anchor,
            context=context,
            name_stack=name_stack,
        )
    )
    accumulator.accept(classification, source_token=value)
    active_function = _active_function_frame(frames)
    if (
        active_function is not None
        and active_function.intrinsic_name == "CHOOSE"
        and active_function.arg_index >= 1
    ):
        if _classification_returns_reference(
            value,
            classification,
            anchor,
            context,
        ):
            active_function.choose_current_reference = True
            active_function.choose_union_pending = False
        else:
            active_function.choose_current_other = True
            active_function.choose_union_pending = False


def _finish_choose_argument(frame: _Frame) -> None:
    if frame.intrinsic_name != "CHOOSE":
        return
    if (
        frame.arg_index >= 1
        and frame.choose_current_reference
        and not frame.choose_current_other
        and not frame.choose_union_pending
        and not frame.choose_union_invalid
    ):
        frame.choose_branch_reference = True
    frame.choose_current_reference = False
    frame.choose_current_other = False
    frame.choose_union_pending = False
    frame.choose_union_invalid = False


def _classification_returns_reference(
    value: str,
    classification: ReferenceClassification,
    anchor: FormulaAnchor,
    context: ReferenceContext,
) -> bool:
    if classification.non_range_name:
        return classification.computed_returns_reference
    candidate = value[:-1] if value.endswith("#") else value
    if candidate.startswith("@"):
        candidate = candidate[1:]
    defined_name = context.defined_name(candidate, anchor.sheet_order)
    if defined_name is not None:
        return defined_name.kind in {"range", "multi_range"}
    return any(
        reference.geometry is not None or reference.via.startswith("external:")
        for reference in classification.references
    )


def _active_function_frame(frames: list[_Frame]) -> _Frame | None:
    return next((frame for frame in reversed(frames) if frame.token_type == "FUNC"), None)


def _matching_groups(tokens: tuple[FormulaToken, ...]) -> dict[int, int]:
    matching: dict[int, int] = {}
    stack: list[int] = []
    for index, token in enumerate(tokens):
        if token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "OPEN":
            stack.append(index)
        elif token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "CLOSE":
            if not stack:  # pragma: no cover - tokenize_formula validates structure
                continue
            opening = stack.pop()
            matching[opening] = index
            matching[index] = opening
    return matching


@dataclass(frozen=True, slots=True)
class _StaticReferenceAtom:
    text: str
    start: int
    end: int


def _fold_static_range_operators(
    tokens: tuple[FormulaToken, ...],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> tuple[FormulaToken, ...]:
    """Collapse transparent static ``left:right`` token groups.

    openpyxl's tokenizer may split the same Excel range operator into several
    shapes depending on whitespace and parentheses.  Folding only operands
    that independently classify as references lets every spelling take the
    same exact range-classification path without guessing about scalar
    expressions or dynamic function results.
    """
    folded = tokens
    while True:
        matching = _matching_groups(folded)
        replacement: tuple[int, int, FormulaToken] | None = None
        for index, token in enumerate(folded):
            if _is_synthetic_range_group_open(token):
                prefix = _function_reference_prefix(token.value)
                closing = matching.get(index)
                if prefix is not None and closing is not None:
                    following = _next_significant(folded, closing)
                    right_start = _next_significant(folded, index)
                    if (
                        right_start is not None
                        and right_start < closing
                        and not (
                            following is not None
                            and following == closing + 1
                            and folded[following].type == "PAREN"
                            and folded[following].subtype == "OPEN"
                        )
                    ):
                        right = _static_reference_atom(
                            folded,
                            matching,
                            right_start,
                            anchor,
                            context,
                            name_stack,
                        )
                        right_end = _previous_significant(folded, closing)
                        if (
                            right is not None
                            and right_end is not None
                            and right.end == right_end
                            and _value_is_composite_endpoint(
                                prefix,
                                anchor,
                                context,
                                name_stack,
                            )
                        ):
                            replacement = (
                                index,
                                closing,
                                FormulaToken(
                                    f"{prefix}:{right.text}",
                                    "OPERAND",
                                    "RANGE",
                                ),
                            )
                            break

            left = _static_reference_atom(
                folded,
                matching,
                index,
                anchor,
                context,
                name_stack,
            )
            if left is None or left.start != index:
                continue
            operator_index = _next_significant(folded, left.end)
            if operator_index is None:
                continue
            if operator_index > left.end + 1 and all(
                folded[space].type == WHITESPACE_TOKEN_TYPE
                for space in range(left.end + 1, operator_index)
            ):
                right = _static_reference_atom(
                    folded,
                    matching,
                    operator_index,
                    anchor,
                    context,
                    name_stack,
                )
                if right is not None:
                    intersection = _static_intersection_token(
                        left.text,
                        right.text,
                        anchor,
                        context,
                        name_stack,
                    )
                    if intersection is not None:
                        replacement = (
                            left.start,
                            right.end,
                            FormulaToken(intersection, "OPERAND", "RANGE"),
                        )
                        break
            operator = folded[operator_index]
            right: _StaticReferenceAtom | None = None
            replacement_end: int | None = None
            if (
                operator.type == "OPERAND"
                and operator.subtype == "RANGE"
                and operator.value.startswith(":")
                and operator.value != ":"
            ):
                right_text = operator.value[1:]
                if _value_is_composite_endpoint(
                    right_text,
                    anchor,
                    context,
                    name_stack,
                ):
                    right = _StaticReferenceAtom(
                        right_text,
                        operator_index,
                        operator_index,
                    )
                    replacement_end = operator_index
            elif (
                operator.type == "OPERAND" and operator.subtype == "RANGE" and operator.value == ":"
            ):
                right_start = _next_significant(folded, operator_index)
                if right_start is not None:
                    right = _static_reference_atom(
                        folded,
                        matching,
                        right_start,
                        anchor,
                        context,
                        name_stack,
                    )
                    if right is not None:
                        replacement_end = right.end
            elif _is_synthetic_range_group_open(operator):
                prefix = _function_reference_prefix(operator.value)
                closing = matching.get(operator_index)
                if prefix is None and closing is not None:
                    right_start = _next_significant(folded, operator_index)
                    if right_start is not None and right_start < closing:
                        candidate = _static_reference_atom(
                            folded,
                            matching,
                            right_start,
                            anchor,
                            context,
                            name_stack,
                        )
                        right_end = _previous_significant(folded, closing)
                        if (
                            candidate is not None
                            and right_end is not None
                            and candidate.end == right_end
                        ):
                            right = candidate
                            replacement_end = closing
            if right is None or replacement_end is None:
                continue
            replacement = (
                left.start,
                replacement_end,
                FormulaToken(
                    f"{left.text}:{right.text}",
                    "OPERAND",
                    "RANGE",
                ),
            )
            break
        if replacement is None:
            return folded
        start, end, token = replacement
        folded = (*folded[:start], token, *folded[end + 1 :])


def _static_reference_atom(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    start: int,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> _StaticReferenceAtom | None:
    if start < 0 or start >= len(tokens):
        return None
    token = tokens[start]
    if token.type == "OPERAND" and token.subtype == "RANGE":
        if token.value == ":" or token.value.startswith(":"):
            return None
        if _value_is_composite_endpoint(token.value, anchor, context, name_stack):
            return _StaticReferenceAtom(token.value, start, start)
        return None
    if token.type == "OPERATOR-PREFIX" and token.value == "@":
        inner_start = _next_significant(tokens, start)
        if inner_start is None:
            return None
        inner = _static_reference_atom(
            tokens,
            matching,
            inner_start,
            anchor,
            context,
            name_stack,
        )
        if inner is None:
            return None
        return _StaticReferenceAtom(f"@{inner.text}", start, inner.end)
    if token.type != "PAREN" or token.subtype != "OPEN":
        return None
    previous = _previous_significant(tokens, start)
    if previous is not None and previous == start - 1 and tokens[previous].subtype == "CLOSE":
        # A group immediately following a closed callable is an invocation
        # argument list, not transparent reference grouping.
        return None
    closing = matching.get(start)
    if closing is None:
        return None
    following = _next_significant(tokens, closing)
    if (
        following is not None
        and following == closing + 1
        and tokens[following].type == "PAREN"
        and tokens[following].subtype == "OPEN"
    ):
        # A transparent group immediately followed by an argument list is a
        # callable expression. Do not consume it as the right side of a range
        # before reference-result inference sees the invocation.
        return None
    inner_start = _next_significant(tokens, start)
    if inner_start is None or inner_start >= closing:
        return None
    inner = _static_reference_atom(
        tokens,
        matching,
        inner_start,
        anchor,
        context,
        name_stack,
    )
    inner_end = _previous_significant(tokens, closing)
    if inner is None or inner_end is None or inner.end != inner_end:
        return None
    return _StaticReferenceAtom(inner.text, start, closing)


def _static_intersection_token(
    left: str,
    right: str,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> str | None:
    classifications = tuple(
        classify_ref(
            value,
            anchor=anchor,
            context=context,
            name_stack=name_stack,
        )
        for value in (left, right)
    )
    if any(
        classification.opaque
        or classification.non_range_name
        or len(classification.references) != 1
        or classification.references[0].geometry is None
        or classification.references[0].via != "ref"
        for classification in classifications
    ):
        return None
    left_reference = classifications[0].references[0]
    right_reference = classifications[1].references[0]
    if (
        left_reference.dst_sheet_order is None
        or left_reference.dst_sheet_order != right_reference.dst_sheet_order
    ):
        return None
    assert left_reference.geometry is not None
    assert right_reference.geometry is not None
    geometry = _intersect_reference_geometries(
        left_reference.geometry,
        right_reference.geometry,
    )
    if geometry is None:
        return None
    local = render_a1_reference(
        ParsedA1Reference("", geometry),
        anchor.cell,
        preserve_range=True,
    )
    if left_reference.dst_sheet_order == anchor.sheet_order:
        return local
    sheet_name = left_reference.dst_sheet_name
    if sheet_name is None:
        return None
    return f"'{sheet_name.replace(chr(39), chr(39) * 2)}'!{local}"


def _unfolded_intersection_metadata(
    tokens: tuple[FormulaToken, ...],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
    *,
    skip_indexes: frozenset[int] = frozenset(),
) -> tuple[tuple[tuple[str, str], ...], bool] | None:
    matching = _matching_groups(tokens)
    dynamic_events: list[tuple[str, str]] = []
    needs_generic_marker = False
    found_intersection = False
    for index, token in enumerate(tokens):
        if token.type != WHITESPACE_TOKEN_TYPE or index in skip_indexes:
            continue
        left_end = _previous_significant(tokens, index)
        right_start = _next_significant(tokens, index)
        if left_end is None or right_start is None:
            continue
        right = _static_reference_atom(
            tokens,
            matching,
            right_start,
            anchor,
            context,
            name_stack,
        )
        if right is None:
            continue
        for left_start in range(left_end, -1, -1):
            left = _static_reference_atom(
                tokens,
                matching,
                left_start,
                anchor,
                context,
                name_stack,
            )
            if left is not None and left.end == left_end:
                found_intersection = True
                classifications = tuple(
                    classify_ref(
                        atom.text,
                        anchor=anchor,
                        context=context,
                        name_stack=name_stack,
                    )
                    for atom in (left, right)
                )
                labels = tuple(
                    dict.fromkeys(
                        label
                        for classification in classifications
                        for label in classification.dynamic_result_labels
                        if not any(
                            reference.via == f"opaque:{label}"
                            for reference in classification.references
                        )
                    )
                )
                source_token = f"{left.text} {right.text}"
                dynamic_events.extend((source_token, label) for label in labels)
                already_opaque = any(
                    classification.opaque
                    or any(
                        reference.via.startswith("opaque:")
                        for reference in classification.references
                    )
                    for classification in classifications
                )
                needs_generic_marker |= not labels and not already_opaque
                break
    if not found_intersection:
        return None
    return tuple(dynamic_events), needs_generic_marker


def _intersect_reference_geometries(
    left: ReferenceGeometry,
    right: ReferenceGeometry,
) -> ReferenceGeometry | None:
    rows = _intersect_axis_terms(
        left.row_a,
        left.row_b,
        right.row_a,
        right.row_b,
        maximum_axis=1_048_576,
    )
    cols = _intersect_axis_terms(
        left.col_a,
        left.col_b,
        right.col_a,
        right.col_b,
        maximum_axis=16_384,
    )
    if rows is None or cols is None:
        return None
    return ReferenceGeometry(rows[0], rows[1], cols[0], cols[1])


def _intersect_axis_terms(
    left_a: AxisTerm,
    left_b: AxisTerm,
    right_a: AxisTerm,
    right_b: AxisTerm,
    *,
    maximum_axis: int,
) -> tuple[AxisTerm, AxisTerm] | None:
    left_values = sorted((left_a.value, left_b.value))
    right_values = sorted((right_a.value, right_b.value))
    left_full = not left_a.relative and not left_b.relative and left_values == [1, maximum_axis]
    right_full = not right_a.relative and not right_b.relative and right_values == [1, maximum_axis]
    if left_full:
        return right_a, right_b
    if right_full:
        return left_a, left_b
    terms = (left_a, left_b, right_a, right_b)
    if len({term.relative for term in terms}) != 1:
        return None
    left_min, left_max = left_values
    right_min, right_max = right_values
    minimum = max(left_min, right_min)
    maximum = min(left_max, right_max)
    if minimum > maximum:
        return None
    relative = left_a.relative
    return AxisTerm(relative, minimum), AxisTerm(relative, maximum)


def _infer_reference_types(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> tuple[_ReferenceInference, bool]:
    inference = _ReferenceInference()
    returns_reference = False
    if tokens:
        returns_reference = _infer_span_reference_type(
            tokens,
            matching,
            0,
            len(tokens) - 1,
            {},
            anchor,
            context,
            name_stack,
            inference,
        )
    return inference, returns_reference


def _result_dynamic_labels(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    inference: _ReferenceInference,
    reference_uses: list[ReferenceUse],
    returns_reference: bool,
) -> tuple[str, ...]:
    if not returns_reference:
        return ()
    transparent = {"LET", "LAMBDA", "SINGLE"}
    positions = set(inference.reference_functions)
    primary = {
        index
        for index in positions
        if normalize_function_name(tokens[index].value) not in transparent
    }
    candidates = primary or positions
    outermost = {
        index
        for index in candidates
        if not any(
            parent != index and parent < index < matching.get(parent, parent)
            for parent in candidates
        )
    }
    labels = [_callable_label(tokens[index].value) for index in sorted(outermost)]
    labels.extend(label for _index, label in sorted(inference.dynamic_callable_groups.items()))
    nested = [label for use in reference_uses for label in use.classification.dynamic_result_labels]
    if nested and (not labels or all(label in transparent for label in labels)):
        labels = nested
    return tuple(dict.fromkeys(label for label in labels if label))


def _infer_span_reference_type(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    start: int,
    end: int,
    bindings: dict[str, bool | _CallableBinding],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
    inference: _ReferenceInference,
    *,
    lexical_endpoint: bool = False,
) -> bool:
    bounds = _trim_whitespace(tokens, start, end)
    if bounds is None:
        return False
    start, end = _unwrap_parentheses(tokens, matching, *bounds)
    has_atom = False
    all_atoms_reference = True
    operators_valid = True
    expect_operand = True

    def accept_atom(returns_reference: bool) -> None:
        nonlocal has_atom, all_atoms_reference, expect_operand
        has_atom = True
        all_atoms_reference = all_atoms_reference and returns_reference
        expect_operand = False

    def accept_reference_operator() -> None:
        nonlocal operators_valid, expect_operand
        if expect_operand:
            operators_valid = False
        expect_operand = True

    index = start
    while index <= end:
        token = tokens[index]
        if token.type == WHITESPACE_TOKEN_TYPE:
            index += 1
            continue
        if token.type == "PAREN" and token.subtype == "OPEN":
            closing = matching.get(index)
            if closing is None or closing > end:
                operators_valid = False
                index += 1
                continue
            group_lexical = lexical_endpoint or _typed_group_is_lexical_colon_endpoint(
                tokens,
                matching,
                index,
                bindings,
            )
            callable_binding = _infer_optional_span_binding_type(
                tokens,
                matching,
                _trim_whitespace(tokens, index + 1, closing - 1),
                bindings,
                anchor,
                context,
                name_stack,
                inference,
                lexical_endpoint=group_lexical,
            )
            invocation_open = _next_significant(tokens, closing)
            if (
                isinstance(callable_binding, _CallableBinding)
                and invocation_open is not None
                and invocation_open <= end
                and tokens[invocation_open].type == "PAREN"
                and tokens[invocation_open].subtype == "OPEN"
            ):
                invocation_close = matching.get(invocation_open)
                if invocation_close is None or invocation_close > end:
                    operators_valid = False
                    index = closing + 1
                    continue
                invocation_arguments = _group_argument_bounds(
                    tokens,
                    invocation_open,
                    invocation_close,
                )
                argument_bindings = tuple(
                    _infer_optional_span_binding_type(
                        tokens,
                        matching,
                        argument,
                        bindings,
                        anchor,
                        context,
                        name_stack,
                        inference,
                    )
                    for argument in invocation_arguments
                )
                result = _apply_callable_binding(
                    callable_binding,
                    argument_bindings,
                    anchor,
                    context,
                    name_stack,
                    inference,
                )
                result, consumed = _consume_lambda_invocations(
                    result,
                    tokens,
                    matching,
                    invocation_close,
                    end,
                    bindings,
                    anchor,
                    context,
                    name_stack,
                    inference,
                )
                if _binding_returns_reference(result):
                    inference.reference_functions.update(
                        _callable_openings(callable_binding, inference)
                    )
                    callable_bounds = _trim_whitespace(tokens, index + 1, closing - 1)
                    if callable_bounds is not None:
                        callable_open, _ = _unwrap_parentheses(
                            tokens,
                            matching,
                            *callable_bounds,
                        )
                        if (
                            tokens[callable_open].type == "FUNC"
                            and tokens[callable_open].subtype == "OPEN"
                        ):
                            inference.reference_functions.add(callable_open)
                        elif _envelope_is_range_endpoint(
                            tokens,
                            index,
                            consumed,
                            False,
                            matching=matching,
                        ):
                            inference.dynamic_callable_groups[index] = _callable_label(
                                tokens[callable_open].value
                            )
                accept_atom(_binding_returns_reference(result))
                index = consumed + 1
                continue
            accept_atom(
                _infer_span_reference_type(
                    tokens,
                    matching,
                    index + 1,
                    closing - 1,
                    bindings,
                    anchor,
                    context,
                    name_stack,
                    inference,
                    lexical_endpoint=group_lexical,
                )
            )
            index = closing + 1
            continue
        if token.type == "ARRAY" and token.subtype == "OPEN":
            closing = matching.get(index)
            if closing is None or closing > end:
                operators_valid = False
                index += 1
                continue
            _infer_span_reference_type(
                tokens,
                matching,
                index + 1,
                closing - 1,
                bindings,
                anchor,
                context,
                name_stack,
                inference,
            )
            accept_atom(False)
            index = closing + 1
            continue
        if _is_synthetic_range_group_open(token):
            prefix = _function_reference_prefix(token.value)
            synthetic_lexical = lexical_endpoint or _typed_synthetic_colon_context(
                tokens,
                matching,
                index,
                prefix,
                bindings,
            )
            if prefix is not None:
                accept_atom(
                    _infer_value_reference_type(
                        prefix,
                        bindings,
                        anchor,
                        context,
                        name_stack,
                        lexical_endpoint=synthetic_lexical,
                    )
                )
            accept_reference_operator()
            closing = matching.get(index)
            if closing is None or closing > end:
                operators_valid = False
                index += 1
                continue
            grouped_type = _infer_optional_span_binding_type(
                tokens,
                matching,
                _trim_whitespace(tokens, index + 1, closing - 1),
                bindings,
                anchor,
                context,
                name_stack,
                inference,
                lexical_endpoint=synthetic_lexical,
            )
            invocation_open = _next_significant(tokens, closing)
            if (
                isinstance(grouped_type, _CallableBinding)
                and invocation_open is not None
                and invocation_open <= end
                and tokens[invocation_open].type == "PAREN"
                and tokens[invocation_open].subtype == "OPEN"
            ):
                invocation_close = matching.get(invocation_open)
                if invocation_close is None or invocation_close > end:
                    operators_valid = False
                    index = closing + 1
                    continue
                invocation_arguments = _group_argument_bounds(
                    tokens,
                    invocation_open,
                    invocation_close,
                )
                argument_bindings = tuple(
                    _infer_optional_span_binding_type(
                        tokens,
                        matching,
                        argument,
                        bindings,
                        anchor,
                        context,
                        name_stack,
                        inference,
                    )
                    for argument in invocation_arguments
                )
                result = _apply_callable_binding(
                    grouped_type,
                    argument_bindings,
                    anchor,
                    context,
                    name_stack,
                    inference,
                )
                result, consumed = _consume_lambda_invocations(
                    result,
                    tokens,
                    matching,
                    invocation_close,
                    end,
                    bindings,
                    anchor,
                    context,
                    name_stack,
                    inference,
                )
                if _binding_returns_reference(result):
                    callable_bounds = _trim_whitespace(tokens, index + 1, closing - 1)
                    if callable_bounds is not None:
                        callable_open, _ = _unwrap_parentheses(
                            tokens,
                            matching,
                            *callable_bounds,
                        )
                        if (
                            tokens[callable_open].type == "FUNC"
                            and tokens[callable_open].subtype == "OPEN"
                        ):
                            inference.reference_functions.add(callable_open)
                        else:
                            inference.dynamic_callable_groups[index] = _callable_label(
                                tokens[callable_open].value
                            )
                accept_atom(_binding_returns_reference(result))
                index = consumed + 1
                continue
            accept_atom(_binding_returns_reference(grouped_type))
            index = closing + 1
            continue
        if token.type == "FUNC" and token.subtype == "OPEN":
            prefix = _function_reference_prefix(token.value)
            if prefix is not None:
                accept_atom(
                    _infer_value_reference_type(
                        prefix,
                        bindings,
                        anchor,
                        context,
                        name_stack,
                    )
                )
                accept_reference_operator()
            result_type, consumed = _infer_function_reference_type(
                tokens,
                matching,
                index,
                end,
                bindings,
                anchor,
                context,
                name_stack,
                inference,
            )
            accept_atom(_binding_returns_reference(result_type))
            index = consumed + 1
            continue
        if token.type == "OPERAND":
            if token.subtype == "RANGE":
                value = token.value
                if value == ":":
                    accept_reference_operator()
                    index += 1
                    continue
                if value.startswith(":"):
                    accept_reference_operator()
                    value = value[1:]
                value_is_lexical_endpoint = lexical_endpoint or _typed_suffix_colon_context(
                    tokens,
                    matching,
                    index,
                    bindings,
                )
                accept_atom(
                    _infer_value_reference_type(
                        value,
                        bindings,
                        anchor,
                        context,
                        name_stack,
                        lexical_endpoint=value_is_lexical_endpoint,
                    )
                )
            else:
                accept_atom(False)
            index += 1
            continue
        if token.type == "OPERATOR-INFIX" and token.value == ",":
            accept_reference_operator()
            index += 1
            continue
        if token.type == "OPERATOR-PREFIX" and token.value == "@":
            # Implicit intersection preserves reference identity when the
            # result is consumed by another reference operator (for example
            # ``CHOOSE(1,@(A1:B5),10):D5``).
            index += 1
            continue
        if token.type.startswith("OPERATOR"):
            operators_valid = False
            expect_operand = True
            index += 1
            continue
        operators_valid = False
        index += 1

    return has_atom and all_atoms_reference and operators_valid and not expect_operand


def _infer_function_reference_type(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    opening: int,
    parent_end: int,
    bindings: dict[str, bool | _CallableBinding],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
    inference: _ReferenceInference,
) -> tuple[bool | _CallableBinding, int]:
    closing = matching.get(opening)
    if closing is None:
        return False, opening
    name = normalize_function_name(tokens[opening].value)
    arguments = _argument_bounds(tokens, opening, closing)

    if name == "LET":
        local_bindings = dict(bindings)
        for argument_index in range(0, max(0, len(arguments) - 1), 2):
            declaration = _binding_name(tokens, arguments[argument_index])
            value_index = argument_index + 1
            value_type = (
                False
                if value_index >= len(arguments) - 1
                else _infer_optional_span_binding_type(
                    tokens,
                    matching,
                    arguments[value_index],
                    local_bindings,
                    anchor,
                    context,
                    name_stack,
                    inference,
                )
            )
            if declaration is not None:
                local_bindings[declaration] = value_type
        result = _infer_optional_span_binding_type(
            tokens,
            matching,
            arguments[-1] if arguments else None,
            local_bindings,
            anchor,
            context,
            name_stack,
            inference,
        )
        result, consumed = _consume_lambda_invocations(
            result,
            tokens,
            matching,
            closing,
            parent_end,
            bindings,
            anchor,
            context,
            name_stack,
            inference,
        )
        if _binding_returns_reference(result):
            inference.reference_functions.add(opening)
        return result, consumed

    if name == "LAMBDA":
        if not arguments:
            return False, closing
        result: bool | _CallableBinding = _LambdaBinding(
            tokens=tokens,
            matching=matching,
            opening=opening,
            parameter_bounds=arguments[:-1],
            body=arguments[-1],
            closure=dict(bindings),
            name_stack=name_stack,
            inference=inference,
        )
        result, consumed = _consume_lambda_invocations(
            result,
            tokens,
            matching,
            closing,
            parent_end,
            bindings,
            anchor,
            context,
            name_stack,
            inference,
        )
        if consumed > closing and _binding_returns_reference(result):
            inference.reference_functions.add(opening)
        return result, consumed

    argument_bindings = tuple(
        _infer_optional_span_binding_type(
            tokens,
            matching,
            argument,
            bindings,
            anchor,
            context,
            name_stack,
            inference,
        )
        for argument in arguments
    )
    argument_types = tuple(_binding_returns_reference(value) for value in argument_bindings)
    if name in ALWAYS_DYNAMIC_REFERENCE_FUNCTIONS:
        inference.reference_functions.add(opening)
        return True, closing
    if name == "INDEX":
        result = bool(argument_types and argument_types[0])
        if result:
            inference.reference_functions.add(opening)
        return result, closing
    if name == "CHOOSE":
        result = any(argument_types[1:])
        if result:
            inference.reference_functions.add(opening)
            inference.dynamic_choices.add(opening)
            return result, closing
        callable_branches = tuple(
            value for value in argument_bindings[1:] if isinstance(value, _CallableBinding)
        )
        if callable_branches:
            return _make_callable_choice(callable_branches), closing
        return False, closing
    if name == "IF":
        callable_branches = tuple(
            value for value in argument_bindings[1:] if isinstance(value, _CallableBinding)
        )
        if callable_branches:
            return _make_callable_choice(callable_branches), closing
        return False, closing
    if name == "SINGLE":
        result = bool(argument_types and argument_types[0])
        if result:
            inference.reference_functions.add(opening)
        return result, closing
    if name in BUILTIN_FUNCTIONS:
        return False, closing
    callable_identifier = compatibility_function_identifier(tokens[opening].value)
    bound_callable = bindings.get(_local_binding_key(callable_identifier))
    if isinstance(bound_callable, _CallableBinding):
        result = _apply_callable_binding(
            bound_callable,
            argument_bindings,
            anchor,
            context,
            name_stack,
            inference,
        )
        result, consumed = _consume_lambda_invocations(
            result,
            tokens,
            matching,
            closing,
            parent_end,
            bindings,
            anchor,
            context,
            name_stack,
            inference,
        )
        if _binding_returns_reference(result):
            inference.reference_functions.add(opening)
        return result, consumed
    defined_name = context.defined_name(callable_identifier, anchor.sheet_order)
    if defined_name is not None and defined_name.kind == "lambda":
        result = _infer_defined_lambda_result_type(
            defined_name,
            argument_bindings,
            anchor,
            context,
            name_stack,
        )
        result, consumed = _consume_lambda_invocations(
            result,
            tokens,
            matching,
            closing,
            parent_end,
            bindings,
            anchor,
            context,
            name_stack,
            inference,
        )
        if _binding_returns_reference(result):
            inference.reference_functions.add(opening)
        return result, consumed
    return False, closing


def _infer_defined_lambda_result_type(
    defined_name: DefinedName,
    argument_bindings: tuple[bool | _CallableBinding, ...],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> bool | _CallableBinding:
    binding = _defined_lambda_binding(
        defined_name,
        anchor,
        context,
        name_stack,
    )
    if binding is None:
        return False
    return _apply_callable_binding(
        binding,
        argument_bindings,
        anchor,
        context,
        name_stack,
        binding.inference,
    )


def _defined_lambda_binding(
    defined_name: DefinedName,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> _LambdaBinding | None:
    key = (defined_name.scope_sheet_order, defined_name.name.casefold())
    if key in name_stack:
        return None
    try:
        tokens = tokenize_formula(defined_name.refers_to)
    except FormulaSyntaxError:
        return None
    matching = _matching_groups(tokens)
    opening = next(
        (index for index, token in enumerate(tokens) if token.type != WHITESPACE_TOKEN_TYPE),
        None,
    )
    if (
        opening is None
        or tokens[opening].type != "FUNC"
        or tokens[opening].subtype != "OPEN"
        or normalize_function_name(tokens[opening].value) != "LAMBDA"
    ):
        return None
    closing = matching.get(opening)
    if closing is None:
        return None
    arguments = _argument_bounds(tokens, opening, closing)
    if not arguments:
        return None
    return _LambdaBinding(
        tokens=tokens,
        matching=matching,
        opening=opening,
        parameter_bounds=arguments[:-1],
        body=arguments[-1],
        closure={},
        name_stack=(*name_stack, key),
        inference=_ReferenceInference(),
    )


def _infer_optional_span_binding_type(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    bounds: tuple[int, int] | None,
    bindings: dict[str, bool | _CallableBinding],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
    inference: _ReferenceInference,
    *,
    lexical_endpoint: bool = False,
) -> bool | _CallableBinding:
    """Infer a value while retaining callable LAMBDA identity."""
    if bounds is None:
        return False
    trimmed = _trim_whitespace(tokens, *bounds)
    if trimmed is None:
        return False
    lambda_binding = _lambda_binding_for_span(
        tokens,
        matching,
        *trimmed,
        bindings,
        name_stack,
        inference,
    )
    if lambda_binding is not None:
        return lambda_binding
    start, end = _unwrap_parentheses(tokens, matching, *trimmed)
    if start == end:
        token = tokens[start]
        if token.type == "OPERAND" and token.subtype == "RANGE":
            candidate = token.value
            if candidate.startswith("@"):
                candidate = candidate[1:]
            if candidate.endswith("#"):
                candidate = candidate[:-1]
            defined_name = context.defined_name(candidate, anchor.sheet_order)
            if defined_name is not None and defined_name.kind == "lambda":
                defined_binding = _defined_lambda_binding(
                    defined_name,
                    anchor,
                    context,
                    name_stack,
                )
                if defined_binding is not None:
                    return defined_binding
    if (
        start <= end
        and tokens[start].type == "FUNC"
        and tokens[start].subtype == "OPEN"
        and not _is_synthetic_range_group_open(tokens[start])
        and _function_reference_prefix(tokens[start].value) is None
    ):
        result, consumed = _infer_function_reference_type(
            tokens,
            matching,
            start,
            end,
            bindings,
            anchor,
            context,
            name_stack,
            inference,
        )
        if consumed == end:
            return result
    return _infer_span_reference_type(
        tokens,
        matching,
        *trimmed,
        bindings,
        anchor,
        context,
        name_stack,
        inference,
        lexical_endpoint=lexical_endpoint,
    )


def _lambda_binding_for_span(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    start: int,
    end: int,
    bindings: dict[str, bool | _CallableBinding],
    name_stack: tuple[NameKey, ...],
    inference: _ReferenceInference,
) -> _CallableBinding | None:
    bounds = _trim_whitespace(tokens, start, end)
    if bounds is None:
        return None
    start, end = _unwrap_parentheses(tokens, matching, *bounds)
    if start == end:
        token = tokens[start]
        if token.type == "OPERAND" and token.subtype == "RANGE":
            candidate = token.value
            if candidate.startswith("@"):
                candidate = candidate[1:]
            if candidate.endswith("#"):
                candidate = candidate[:-1]
            bound = bindings.get(_local_binding_key(candidate))
            return bound if isinstance(bound, _CallableBinding) else None
        return None
    opening = tokens[start]
    if (
        opening.type != "FUNC"
        or opening.subtype != "OPEN"
        or normalize_function_name(opening.value) != "LAMBDA"
        or matching.get(start) != end
    ):
        return None
    arguments = _argument_bounds(tokens, start, end)
    if not arguments:
        return None
    return _LambdaBinding(
        tokens=tokens,
        matching=matching,
        opening=start,
        parameter_bounds=arguments[:-1],
        body=arguments[-1],
        closure=dict(bindings),
        name_stack=name_stack,
        inference=inference,
    )


def _apply_callable_binding(
    binding: _CallableBinding,
    argument_bindings: tuple[bool | _CallableBinding, ...],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
    inference: _ReferenceInference,
) -> bool | _CallableBinding:
    if isinstance(binding, _CallableChoice):
        if binding.truncated:
            # A discarded alternative may be reference-returning.  Once the
            # selector is invoked, visible opacity is safer than silently
            # treating that unknown result as scalar.
            return True
        results = tuple(
            _apply_callable_binding(
                alternative,
                argument_bindings,
                anchor,
                context,
                name_stack,
                inference,
            )
            for alternative in binding.alternatives
        )
        if any(_binding_returns_reference(result) for result in results):
            return True
        callable_results = tuple(
            result for result in results if isinstance(result, _CallableBinding)
        )
        if not callable_results:
            return False
        return _make_callable_choice(callable_results)
    assert isinstance(binding, _LambdaBinding)
    key = (id(binding.tokens), binding.opening)
    owner_inference = binding.inference
    if key in owner_inference.active_lambda_calls:
        return False
    owner_inference.active_lambda_calls.add(key)
    try:
        local_bindings = dict(binding.closure)
        for parameter_index, parameter in enumerate(binding.parameter_bounds):
            declaration = _binding_name(binding.tokens, parameter)
            if declaration is not None:
                local_bindings[declaration] = (
                    argument_bindings[parameter_index]
                    if parameter_index < len(argument_bindings)
                    else False
                )
        combined_name_stack = tuple(dict.fromkeys((*name_stack, *binding.name_stack)))
        return _infer_optional_span_binding_type(
            binding.tokens,
            binding.matching,
            binding.body,
            local_bindings,
            anchor,
            context,
            combined_name_stack,
            owner_inference,
        )
    finally:
        owner_inference.active_lambda_calls.remove(key)


def _make_callable_choice(
    alternatives: tuple[_CallableBinding, ...],
) -> _CallableBinding:
    """Flatten, deduplicate, and bound callable selector alternatives.

    Applying nested selectors can otherwise replicate the same choice tree at
    every invocation and grow exponentially.  Encounter order is stable, and
    exact object identity is sufficient to collapse the repeated branches
    produced by applying one lexical choice to another.  Distinct closures
    remain distinct until the fixed cap; overflow is retained as an explicit
    conservative flag rather than being silently discarded.
    """
    normalized: list[_CallableBinding] = []
    truncated = False

    def append(binding: _CallableBinding) -> None:
        nonlocal truncated
        if isinstance(binding, _CallableChoice):
            truncated = truncated or binding.truncated
            for alternative in binding.alternatives:
                append(alternative)
            return
        if any(binding is existing for existing in normalized):
            return
        if len(normalized) >= _MAX_CALLABLE_CHOICE_ALTERNATIVES:
            truncated = True
            return
        normalized.append(binding)

    for alternative in alternatives:
        append(alternative)

    # Callers only construct a choice from at least one callable branch.
    assert normalized
    if len(normalized) == 1 and not truncated:
        return normalized[0]
    return _CallableChoice(tuple(normalized), truncated=truncated)


def _consume_lambda_invocations(
    result: bool | _CallableBinding,
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    consumed: int,
    parent_end: int,
    bindings: dict[str, bool | _CallableBinding],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
    inference: _ReferenceInference,
) -> tuple[bool | _CallableBinding, int]:
    while isinstance(result, _CallableBinding):
        invocation_open = _next_significant(tokens, consumed)
        if (
            invocation_open is None
            or invocation_open > parent_end
            or tokens[invocation_open].type != "PAREN"
            or tokens[invocation_open].subtype != "OPEN"
        ):
            break
        invocation_close = matching.get(invocation_open)
        if invocation_close is None or invocation_close > parent_end:
            break
        invocation_arguments = _group_argument_bounds(
            tokens,
            invocation_open,
            invocation_close,
        )
        argument_bindings = tuple(
            _infer_optional_span_binding_type(
                tokens,
                matching,
                argument,
                bindings,
                anchor,
                context,
                name_stack,
                inference,
            )
            for argument in invocation_arguments
        )
        result = _apply_callable_binding(
            result,
            argument_bindings,
            anchor,
            context,
            name_stack,
            inference,
        )
        consumed = invocation_close
    return result, consumed


def _binding_returns_reference(value: bool | _CallableBinding) -> bool:
    return value is True


def _callable_openings(
    binding: _CallableBinding,
    inference: _ReferenceInference,
) -> set[int]:
    if isinstance(binding, _LambdaBinding):
        return {binding.opening} if binding.inference is inference else set()
    assert isinstance(binding, _CallableChoice)
    return {
        opening
        for alternative in binding.alternatives
        for opening in _callable_openings(alternative, inference)
    }


def _infer_value_reference_type(
    value: str,
    bindings: dict[str, bool | _CallableBinding],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
    *,
    lexical_endpoint: bool = False,
) -> bool:
    composite_type = _lexical_composite_reference_type(
        value,
        bindings,
        anchor,
        context,
        name_stack,
    )
    if composite_type is not None:
        return composite_type
    candidate = value
    if candidate.startswith("@"):
        candidate = candidate[1:]
    if candidate.endswith("#"):
        candidate = candidate[:-1]
    bound_type = bindings.get(_local_binding_key(candidate))
    if bound_type is not None:
        return _binding_returns_reference(bound_type)
    classification = (
        _classify_lexical_endpoint(
            value,
            anchor=anchor,
            context=context,
            name_stack=name_stack,
        )
        if lexical_endpoint
        else classify_ref(
            value,
            anchor=anchor,
            context=context,
            name_stack=name_stack,
        )
    )
    return _classification_returns_reference(
        value,
        classification,
        anchor,
        context,
    )


def _binding_name(
    tokens: tuple[FormulaToken, ...],
    bounds: tuple[int, int] | None,
) -> str | None:
    if bounds is None:
        return None
    start, end = bounds
    if start != end:
        return None
    token = tokens[start]
    if (
        token.type != "OPERAND"
        or token.subtype != "RANGE"
        or not _is_declaration_identifier(token.value)
    ):
        return None
    return _local_binding_key(token.value)


def _group_argument_bounds(
    tokens: tuple[FormulaToken, ...],
    opening: int,
    closing: int,
) -> tuple[tuple[int, int] | None, ...]:
    start = opening + 1
    depth = 0
    arguments: list[tuple[int, int] | None] = []
    for index in range(start, closing):
        token = tokens[index]
        if token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "OPEN":
            depth += 1
        elif token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "CLOSE":
            depth -= 1
        elif token.type == "OPERATOR-INFIX" and token.value == "," and depth == 0:
            arguments.append(_trim_whitespace(tokens, start, index - 1))
            start = index + 1
    arguments.append(_trim_whitespace(tokens, start, closing - 1))
    return tuple(arguments)


def _argument_bounds(
    tokens: tuple[FormulaToken, ...],
    opening: int,
    closing: int,
) -> tuple[tuple[int, int] | None, ...]:
    start = opening + 1
    depth = 0
    arguments: list[tuple[int, int] | None] = []
    for index in range(start, closing):
        token = tokens[index]
        if token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "OPEN":
            depth += 1
        elif token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "CLOSE":
            depth -= 1
        elif token.type == "SEP" and token.subtype == "ARG" and depth == 0:
            arguments.append(_trim_whitespace(tokens, start, index - 1))
            start = index + 1
    arguments.append(_trim_whitespace(tokens, start, closing - 1))
    return tuple(arguments)


def _trim_whitespace(
    tokens: tuple[FormulaToken, ...],
    start: int,
    end: int,
) -> tuple[int, int] | None:
    while start <= end and tokens[start].type == WHITESPACE_TOKEN_TYPE:
        start += 1
    while start <= end and tokens[end].type == WHITESPACE_TOKEN_TYPE:
        end -= 1
    return None if start > end else (start, end)


def _unwrap_parentheses(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    start: int,
    end: int,
) -> tuple[int, int]:
    while True:
        bounds = _trim_whitespace(tokens, start, end)
        if bounds is None:
            return start, end
        start, end = bounds
        if not (
            start < end
            and tokens[start].type == "PAREN"
            and tokens[start].subtype == "OPEN"
            and matching.get(start) == end
        ):
            return start, end
        start += 1
        end -= 1
    return start, end


def _value_is_composite_endpoint(
    value: str,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> bool:
    """Return whether a value must be retained as a composite endpoint.

    Formula/LAMBDA names are candidates syntactically, but their body
    precedents are not exact result geometry. Folding keeps the operator
    visible to ``classify_ref``, which then marks it conservative.
    """
    classification = classify_ref(
        value,
        anchor=anchor,
        context=context,
        name_stack=name_stack,
    )
    return (
        classification.non_range_name
        or any(reference.via == "opaque:structured" for reference in classification.references)
        or _classification_returns_reference(
            value,
            classification,
            anchor,
            context,
        )
    )


def _dynamic_index_positions(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    reference_indexes: frozenset[int],
    reference_functions: frozenset[int] = frozenset(),
) -> frozenset[int]:
    dynamic: set[int] = set()
    parents = _group_parents(tokens)
    for index in reference_indexes:
        envelope = _reference_function_envelope(
            tokens,
            matching,
            index,
        )
        if envelope is None:
            continue
        left, right, colon_before = envelope
        parent = parents.get(left)
        while (
            parent is not None
            and parent in reference_functions
            and normalize_function_name(tokens[parent].value) in {"LET", "LAMBDA", "SINGLE"}
        ):
            parent_envelope = _reference_function_envelope(
                tokens,
                matching,
                parent,
            )
            if parent_envelope is None:
                break
            left, right, parent_colon_before = parent_envelope
            colon_before = colon_before or parent_colon_before
            parent = parents.get(left)

        if _envelope_is_range_endpoint(tokens, left, right, colon_before):
            dynamic.add(index)
    return frozenset(dynamic)


def _reference_function_range_endpoint_positions(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    reference_functions: frozenset[int],
) -> frozenset[int]:
    endpoints: set[int] = set()
    for index in reference_functions:
        envelope = _reference_function_envelope(tokens, matching, index)
        if envelope is None:
            continue
        left, right, colon_before = envelope
        if _envelope_is_range_endpoint(tokens, left, right, colon_before):
            endpoints.add(index)
    return frozenset(endpoints)


def _reference_function_envelope(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    index: int,
) -> tuple[int, int, bool] | None:
    closing = matching.get(index)
    if closing is None:
        return None
    left = index
    right = closing
    colon_before = ":" in tokens[index].value[:-1]
    while True:
        changed = False
        wrapper = _previous_significant(tokens, left)
        wrapper_close = _next_significant(tokens, right)
        if (
            wrapper is not None
            and wrapper_close is not None
            and _is_grouping_open(tokens[wrapper])
            and matching.get(wrapper) == wrapper_close
        ):
            colon_before = colon_before or _is_synthetic_range_group_open(tokens[wrapper])
            left = wrapper
            right = wrapper_close
            changed = True
        prefix = _previous_significant(tokens, left)
        if (
            prefix is not None
            and tokens[prefix].type == "OPERATOR-PREFIX"
            and tokens[prefix].value == "@"
        ):
            left = prefix
            changed = True
        invocation = _next_significant(tokens, right)
        if (
            invocation is not None
            and tokens[invocation].type == "PAREN"
            and tokens[invocation].subtype == "OPEN"
        ):
            invocation_close = matching.get(invocation)
            if invocation_close is not None:
                right = invocation_close
                changed = True
        if not changed:
            return left, right, colon_before


def _envelope_is_range_endpoint(
    tokens: tuple[FormulaToken, ...],
    left: int,
    right: int,
    colon_before: bool,
    *,
    matching: dict[int, int] | None = None,
) -> bool:
    if matching is not None:
        while True:
            changed = False
            wrapper = _previous_significant(tokens, left)
            wrapper_close = _next_significant(tokens, right)
            if (
                wrapper is not None
                and wrapper_close is not None
                and _is_grouping_open(tokens[wrapper])
                and matching.get(wrapper) == wrapper_close
            ):
                colon_before = colon_before or _is_synthetic_range_group_open(tokens[wrapper])
                left = wrapper
                right = wrapper_close
                changed = True
            prefix = _previous_significant(tokens, left)
            if (
                prefix is not None
                and tokens[prefix].type == "OPERATOR-PREFIX"
                and tokens[prefix].value == "@"
            ):
                left = prefix
                changed = True
            if not changed:
                break
    if colon_before:
        return True
    previous = _previous_significant(tokens, left)
    if previous is not None and tokens[previous].value == ":":
        return True
    following = _next_significant(tokens, right)
    if following is not None and tokens[following].value.startswith(":"):
        return True
    if left > 0 and tokens[left - 1].type == WHITESPACE_TOKEN_TYPE:
        previous = _previous_significant(tokens, left)
        if previous is not None and _token_can_end_intersection_operand(tokens[previous]):
            return True
    if right + 1 < len(tokens) and tokens[right + 1].type == WHITESPACE_TOKEN_TYPE:
        following = _next_significant(tokens, right)
        if following is not None and _token_can_start_intersection_operand(tokens[following]):
            return True
    return False


def _token_can_start_intersection_operand(token: FormulaToken) -> bool:
    return bool(
        (token.type == "OPERAND" and token.subtype == "RANGE")
        or (token.type in {"FUNC", "PAREN"} and token.subtype == "OPEN")
        or (token.type == "OPERATOR-PREFIX" and token.value == "@")
    )


def _token_can_end_intersection_operand(token: FormulaToken) -> bool:
    return bool(
        (token.type == "OPERAND" and token.subtype == "RANGE")
        or (token.type in {"FUNC", "PAREN"} and token.subtype == "CLOSE")
    )


def _group_parents(tokens: tuple[FormulaToken, ...]) -> dict[int, int]:
    parents: dict[int, int] = {}
    stack: list[int] = []
    for index, token in enumerate(tokens):
        if stack:
            parents[index] = stack[-1]
        if token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "OPEN":
            stack.append(index)
        elif token.type in {"FUNC", "PAREN", "ARRAY"} and token.subtype == "CLOSE" and stack:
            stack.pop()
    return parents


def _previous_significant(
    tokens: tuple[FormulaToken, ...],
    index: int,
) -> int | None:
    candidate = index - 1
    while candidate >= 0:
        if tokens[candidate].type != WHITESPACE_TOKEN_TYPE:
            return candidate
        candidate -= 1
    return None


def _next_significant(
    tokens: tuple[FormulaToken, ...],
    index: int,
) -> int | None:
    candidate = index + 1
    while candidate < len(tokens):
        if tokens[candidate].type != WHITESPACE_TOKEN_TYPE:
            return candidate
        candidate += 1
    return None


def _is_grouping_open(token: FormulaToken) -> bool:
    return bool(
        (token.type == "PAREN" and token.subtype == "OPEN") or _is_synthetic_range_group_open(token)
    )


def _is_synthetic_range_group_open(token: FormulaToken) -> bool:
    return bool(
        token.type == "FUNC"
        and token.subtype == "OPEN"
        and token.value[:-1].rstrip().endswith(":")
        and not compatibility_function_identifier(token.value)
    )


def _token_can_start_reference_expression(
    token: FormulaToken,
    *,
    function_returns_reference: bool,
) -> bool:
    if token.type == "OPERATOR-INFIX" and token.value == ",":
        return True
    if token.type == "OPERATOR-PREFIX" and token.value == "@":
        return True
    if token.type == "OPERAND" and token.subtype == "RANGE":
        return True
    if token.type == "PAREN" and token.subtype == "OPEN":
        return True
    if token.type != "FUNC" or token.subtype != "OPEN":
        return False
    if _is_synthetic_range_group_open(token):
        return True
    return function_returns_reference or normalize_function_name(token.value) == "CHOOSE"


def _frame_returns_reference(frame: _Frame) -> bool:
    return frame.returns_reference or bool(
        frame.intrinsic_name == "CHOOSE" and frame.choose_branch_reference
    )


def _argument_can_declare(frame: _Frame) -> bool:
    if frame.intrinsic_name == "LAMBDA":
        return True
    return frame.intrinsic_name == "LET" and frame.arg_index % 2 == 0


def _is_declaration_identifier(value: str) -> bool:
    if not value or value.startswith("@") or value.endswith("#"):
        return False
    candidate = value
    stored_parameter = candidate.casefold().startswith(_LOCAL_PARAMETER_PREFIX)
    if stored_parameter:
        candidate = candidate[len(_LOCAL_PARAMETER_PREFIX) :]
    if not candidate:
        return False
    first = candidate[0]
    if not (first.isalpha() or first in {"_", "\\"}):
        return False
    if not all(character.isalnum() or character in {"_", "\\"} for character in candidate[1:]):
        return False
    return stored_parameter or parse_a1_reference(candidate, CellRef(1, 1)) is None


def _local_binding_key(value: str) -> str:
    return value.casefold()


def _local_display_key(value: str) -> str:
    folded = _local_binding_key(value)
    if folded.startswith(_LOCAL_PARAMETER_PREFIX):
        return folded[len(_LOCAL_PARAMETER_PREFIX) :]
    return folded


def _lexical_composite_endpoints(
    value: str,
    frames: list[_Frame],
    anchor: FormulaAnchor,
) -> tuple[str, str] | None:
    spelling = _lexical_colon_spelling(value)
    if spelling is not None and any(_is_bound(endpoint, frames) for endpoint in spelling):
        return spelling
    for parsed in parse_reference_range_candidates(value, anchor.cell):
        endpoints = (parsed.left.original, parsed.right.original)
        if any(_is_bound(endpoint, frames) for endpoint in endpoints):
            return endpoints
    return None


def _lexical_composite_reference_type(
    value: str,
    bindings: dict[str, bool | _CallableBinding],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> bool | None:
    spelling = _lexical_colon_spelling(value)
    if spelling is not None:
        spelling_type = _lexical_endpoint_reference_type(
            spelling,
            bindings,
            anchor,
            context,
            name_stack,
        )
        if spelling_type is not None:
            return spelling_type
    for parsed in parse_reference_range_candidates(value, anchor.cell):
        parsed_type = _lexical_endpoint_reference_type(
            (parsed.left.original, parsed.right.original),
            bindings,
            anchor,
            context,
            name_stack,
        )
        if parsed_type is not None:
            return parsed_type
    return None


def _lexical_colon_spelling(value: str) -> tuple[str, str] | None:
    if value.count(":") != 1 or any(character in value for character in "![]'"):
        return None
    left, right = value.split(":", 1)
    if not left or not right:
        return None
    return left, right


def _lexical_endpoint_reference_type(
    endpoints: tuple[str, str],
    bindings: dict[str, bool | _CallableBinding],
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> bool | None:
    endpoint_types: list[bool] = []
    bound_seen = False
    for endpoint in endpoints:
        candidate = endpoint
        if candidate.startswith("@"):
            candidate = candidate[1:]
        if candidate.endswith("#"):
            candidate = candidate[:-1]
        binding = bindings.get(_local_binding_key(candidate))
        if binding is not None:
            bound_seen = True
            endpoint_types.append(_binding_returns_reference(binding))
            continue
        classification = _classify_lexical_endpoint(
            endpoint,
            anchor=anchor,
            context=context,
            name_stack=name_stack,
        )
        endpoint_types.append(
            _classification_returns_reference(
                endpoint,
                classification,
                anchor,
                context,
            )
        )
    return all(endpoint_types) if bound_seen else None


def _classify_lexical_endpoint(
    value: str,
    *,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> ReferenceClassification:
    classification = classify_ref(
        value,
        anchor=anchor,
        context=context,
        name_stack=name_stack,
    )
    if classification.non_range_name or not any(
        issue.code == "W_UNKNOWN_NAME" for issue in classification.issues
    ):
        return classification
    candidate = value
    if candidate.startswith("@"):
        candidate = candidate[1:]
    if candidate.endswith("#"):
        candidate = candidate[:-1]
    if not (candidate.isdecimal() or (candidate.isascii() and candidate.isalpha())):
        return classification
    axis = classify_ref(
        f"{candidate}:{candidate}",
        anchor=anchor,
        context=context,
        name_stack=name_stack,
    )
    if any(reference.geometry is not None for reference in axis.references):
        return axis
    return classification


def _lexical_intersection_marker(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    whitespace: int,
    frames: list[_Frame],
) -> str | None:
    left_index = _previous_significant(tokens, whitespace)
    right_index = _next_significant(tokens, whitespace)
    if left_index is None or right_index is None:
        return None
    left = _lexical_intersection_operand(tokens, matching, left_index)
    right = _lexical_intersection_operand(tokens, matching, right_index)
    if left is None or right is None:
        return None
    if not (_is_bound(left, frames) or _is_bound(right, frames)):
        return None
    return f"{left} {right}"


def _group_is_lexical_colon_endpoint(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    opening: int,
    frames: list[_Frame],
) -> bool:
    token = tokens[opening]
    if token.type != "PAREN" or token.subtype != "OPEN":
        return False
    closing = matching.get(opening)
    if closing is None:
        return False
    left = _lexical_intersection_operand(tokens, matching, opening)
    if left is None:
        return False
    following = _next_significant(tokens, closing)
    if following is None:
        return False
    right = tokens[following]
    if right.type == "OPERAND" and right.subtype == "RANGE" and right.value.startswith(":"):
        suffix = right.value[1:]
        return bool(suffix) and (_is_bound(left, frames) or _is_bound(suffix, frames))
    if _is_synthetic_range_group_open(right):
        return (
            _lexical_synthetic_colon_marker(
                tokens,
                matching,
                following,
                _function_reference_prefix(right.value),
                frames,
            )
            is not None
        )
    return False


def _typed_is_bound(
    value: str,
    bindings: dict[str, bool | _CallableBinding],
) -> bool:
    candidate = value
    if candidate.startswith("@"):
        candidate = candidate[1:]
    if candidate.endswith("#"):
        candidate = candidate[:-1]
    return _local_binding_key(candidate) in bindings


def _typed_group_is_lexical_colon_endpoint(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    opening: int,
    bindings: dict[str, bool | _CallableBinding],
) -> bool:
    closing = matching.get(opening)
    if closing is None:
        return False
    left = _lexical_intersection_operand(tokens, matching, opening)
    if left is None:
        return False
    following = _next_significant(tokens, closing)
    if following is None:
        return False
    right = tokens[following]
    if right.type == "OPERAND" and right.subtype == "RANGE" and right.value.startswith(":"):
        suffix = right.value[1:]
        return bool(suffix) and (
            _typed_is_bound(left, bindings) or _typed_is_bound(suffix, bindings)
        )
    if _is_synthetic_range_group_open(right):
        return _typed_synthetic_colon_context(
            tokens,
            matching,
            following,
            _function_reference_prefix(right.value),
            bindings,
        )
    return False


def _typed_suffix_colon_context(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    index: int,
    bindings: dict[str, bool | _CallableBinding],
) -> bool:
    suffix = tokens[index].value
    if not suffix.startswith(":") or len(suffix) == 1:
        return False
    left_index = _previous_significant(tokens, index)
    if left_index is None:
        return False
    left = _lexical_intersection_operand(tokens, matching, left_index)
    return left is not None and (
        _typed_is_bound(left, bindings) or _typed_is_bound(suffix[1:], bindings)
    )


def _typed_synthetic_colon_context(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    opening: int,
    prefix: str | None,
    bindings: dict[str, bool | _CallableBinding],
) -> bool:
    closing = matching.get(opening)
    if closing is None:
        return False
    bounds = _trim_whitespace(tokens, opening + 1, closing - 1)
    if bounds is None:
        return False
    start, end = _unwrap_parentheses(tokens, matching, *bounds)
    if start != end:
        return False
    operand = tokens[start]
    if operand.type != "OPERAND" or operand.subtype != "RANGE":
        return False
    left = prefix
    if not left:
        left_index = _previous_significant(tokens, opening)
        if left_index is None:
            return False
        left = _lexical_intersection_operand(tokens, matching, left_index)
    return left is not None and (
        _typed_is_bound(left, bindings) or _typed_is_bound(operand.value, bindings)
    )


def _lexical_suffix_colon_marker(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    index: int,
    frames: list[_Frame],
) -> str | None:
    suffix = tokens[index].value
    if not suffix.startswith(":") or len(suffix) == 1:
        return None
    left_index = _previous_significant(tokens, index)
    if left_index is None:
        return None
    left = _lexical_intersection_operand(tokens, matching, left_index)
    right = suffix[1:]
    if left is None or not (_is_bound(left, frames) or _is_bound(right, frames)):
        return None
    return f"{left}{suffix}"


def _lexical_synthetic_colon_marker(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    opening: int,
    prefix: str | None,
    frames: list[_Frame],
) -> str | None:
    if not _is_synthetic_range_group_open(tokens[opening]):
        return None
    closing = matching.get(opening)
    if closing is None:
        return None
    bounds = _trim_whitespace(tokens, opening + 1, closing - 1)
    if bounds is None:
        return None
    start, end = _unwrap_parentheses(tokens, matching, *bounds)
    if start != end:
        return None
    operand = tokens[start]
    if operand.type != "OPERAND" or operand.subtype != "RANGE":
        return None
    left = prefix
    if not left:
        left_index = _previous_significant(tokens, opening)
        if left_index is None:
            return None
        left = _lexical_intersection_operand(tokens, matching, left_index)
    if left is None or not (_is_bound(left, frames) or _is_bound(operand.value, frames)):
        return None
    return f"{left}:{operand.value}"


def _lexical_intersection_operand(
    tokens: tuple[FormulaToken, ...],
    matching: dict[int, int],
    adjacent: int,
) -> str | None:
    token = tokens[adjacent]
    if token.type == "PAREN" and token.subtype == "CLOSE":
        opening = matching.get(adjacent)
        if opening is None:
            return None
        start, end = _unwrap_parentheses(tokens, matching, opening, adjacent)
    elif token.type == "PAREN" and token.subtype == "OPEN":
        closing = matching.get(adjacent)
        if closing is None:
            return None
        start, end = _unwrap_parentheses(tokens, matching, adjacent, closing)
    else:
        start = end = adjacent
    if start != end:
        return None
    operand = tokens[start]
    if (
        operand.type != "OPERAND"
        or operand.subtype != "RANGE"
        or operand.value == ":"
        or operand.value.startswith(":")
    ):
        return None
    return operand.value


def _is_bound(value: str, frames: list[_Frame]) -> bool:
    candidate = value
    if candidate.startswith("@"):
        candidate = candidate[1:]
    if candidate.endswith("#"):
        candidate = candidate[:-1]
    folded = _local_binding_key(candidate)
    return any(folded in frame.bindings for frame in reversed(frames) if frame.special)


def _function_reference_prefix(value: str) -> str | None:
    body = value[:-1] if value.endswith("(") else value
    if ":" not in body:
        return None
    prefix = body.rsplit(":", 1)[0].lstrip("@")
    return prefix or None


def _callable_label(value: str) -> str:
    label = normalize_function_name(_local_display_key(value))
    return label or "CALLABLE"


def _emit_dynamic(accumulator: _Accumulator, frame: _Frame) -> None:
    if frame.dynamic_emitted or frame.name is None:
        return
    frame.dynamic_emitted = True
    accumulator.references.append(
        ExtractedReference(frame.source_token, None, None, None, f"opaque:{frame.name}")
    )
    accumulator.issues.append(
        FormulaIssue(
            "info",
            "I_DYNAMIC_REF",
            f"Function {frame.name} returns a dynamic reference.",
            {"function": frame.name},
        )
    )
    accumulator.opaque = True


def _emit_dynamic_label(
    accumulator: _Accumulator,
    *,
    source_token: str,
    label: str,
) -> None:
    accumulator.references.append(
        ExtractedReference(source_token, None, None, None, f"opaque:{label}")
    )
    accumulator.issues.append(
        FormulaIssue(
            "info",
            "I_DYNAMIC_REF",
            f"Callable {label} returns a dynamic reference.",
            {"function": label},
        )
    )
    accumulator.opaque = True


def _mark_parse_problem(accumulator: _Accumulator, token: str, message: str) -> None:
    accumulator.issues.append(FormulaIssue("warn", "W_PARSE", message, {"token": token}))
    if not any(reference.via == "opaque:parse" for reference in accumulator.references):
        accumulator.references.append(ExtractedReference(token, None, None, None, "opaque:parse"))
    accumulator.opaque = True


def _failed_analysis(formula: str, message: str) -> FormulaAnalysis:
    return FormulaAnalysis(
        references=(ExtractedReference(formula, None, None, None, "opaque:parse"),),
        function_calls=(),
        volatile=False,
        opaque=True,
        issues=(FormulaIssue("warn", "W_PARSE", message, {"formula": formula}),),
        reference_uses=(),
    )


__all__ = ["FormulaAnalysis", "ReferenceUse", "analyze_formula"]
