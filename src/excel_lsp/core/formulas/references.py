"""Workbook-aware classification of Excel formula reference operands."""

from __future__ import annotations

import re
from asyncio import current_task
from collections.abc import Generator, Iterable, Mapping
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass, field, replace
from threading import current_thread
from types import MappingProxyType
from typing import Literal

from excel_lsp.core.external_links import external_link_label
from excel_lsp.core.formulas.a1 import (
    AxisTerm,
    CellRef,
    ParsedReferenceRange,
    ReferenceGeometry,
    modern_range_endpoint_geometries,
    parse_a1_reference,
    parse_modern_a1_range,
    parse_reference_range,
    parse_reference_range_candidates,
)
from excel_lsp.core.formulas.tokens import FormulaSyntaxError, tokenize_formula
from excel_lsp.core.models import DefinedName, Rect, SheetDescriptor, TableInfo
from excel_lsp.core.parse.coordinates import parse_rect

IssueSeverity = Literal["info", "warn"]
NameKey = tuple[int | None, str]

_R1C1_REFERENCE = re.compile(
    r"R(?:\d+|\[[-+]?\d+\])?C(?:\d+|\[[-+]?\d+\])?(?::"
    r"R(?:\d+|\[[-+]?\d+\])?C(?:\d+|\[[-+]?\d+\])?)?",
    re.IGNORECASE,
)
_EXTERNAL_BOOK = re.compile(r"\[([^\]]+)\]")


def _empty_related() -> Mapping[str, object]:
    return MappingProxyType({})


def _empty_external_links() -> Mapping[int, str]:
    return MappingProxyType({})


@dataclass(frozen=True, slots=True)
class FormulaAnchor:
    """A formula cell plus its workbook sheet identity."""

    sheet_order: int
    sheet_name: str
    row: int
    col: int

    def __post_init__(self) -> None:
        if type(self.sheet_order) is not int or self.sheet_order < 0:
            raise ValueError("sheet_order must be a nonnegative integer")
        if not self.sheet_name:
            raise ValueError("sheet_name must not be empty")
        CellRef(self.row, self.col)

    @property
    def cell(self) -> CellRef:
        """Return the local coordinate used by the A1 geometry layer."""
        return CellRef(self.row, self.col)


@dataclass(frozen=True, slots=True)
class TableBinding:
    """One ListObject associated with its owning worksheet."""

    sheet_order: int
    sheet_name: str
    table: TableInfo


@dataclass(frozen=True, slots=True)
class _StructuredSelector:
    text: str
    escaped_first: bool = False


@dataclass(frozen=True, slots=True)
class _TableSpatialNode:
    bounds: Rect
    entries: tuple[tuple[TableBinding, Rect], ...] = ()
    left: _TableSpatialNode | None = None
    right: _TableSpatialNode | None = None


@dataclass(frozen=True, slots=True)
class FormulaIssue:
    """One analysis-time diagnostic independent of persistence ids."""

    severity: IssueSeverity
    code: str
    message: str
    related: Mapping[str, object] = field(default_factory=_empty_related)

    def __post_init__(self) -> None:
        object.__setattr__(self, "related", MappingProxyType(dict(self.related)))


@dataclass(frozen=True, slots=True)
class ExtractedReference:
    """One resolved, external, or deliberately opaque formula edge."""

    token: str
    dst_sheet_order: int | None
    dst_sheet_name: str | None
    geometry: ReferenceGeometry | None
    via: str
    # Composite ranges need their endpoint geometries at block-extrusion time;
    # ``geometry`` remains the exact range hull at the analysis anchor.
    extrusion_geometries: tuple[ReferenceGeometry, ...] = ()


@dataclass(frozen=True, slots=True)
class ReferenceClassification:
    """Result of classifying one operand or callable defined name."""

    references: tuple[ExtractedReference, ...] = ()
    issues: tuple[FormulaIssue, ...] = ()
    function_calls: tuple[str, ...] = ()
    volatile: bool = False
    opaque: bool = False
    # Non-range defined names expose body precedents or a scalar constant, not
    # an exact endpoint geometry. Composite operators retain them but must not
    # synthesize a rectangle from those precedents.
    non_range_name: bool = False
    computed_returns_reference: bool = False
    dynamic_result_labels: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True, eq=False)
class _IdentityKey:
    value: object

    def __hash__(self) -> int:
        return id(self.value)

    def __eq__(self, other: object) -> bool:
        return isinstance(other, _IdentityKey) and self.value is other.value


_NameCacheKey = tuple[
    _IdentityKey,
    NameKey,
    int,
    str,
    int,
    int,
    str,
    bool,
    bool,
    tuple[NameKey, ...],
]


def _empty_name_cache() -> dict[_NameCacheKey, ReferenceClassification]:
    return {}


@dataclass(slots=True)
class _NameExpansionState:
    owner: tuple[object, object | None]
    cache: dict[_NameCacheKey, ReferenceClassification] = field(default_factory=_empty_name_cache)


_NAME_EXPANSION_CACHE: ContextVar[_NameExpansionState | None] = ContextVar(
    "excel_lsp_name_expansion_cache",
    default=None,
)


@dataclass(frozen=True, slots=True)
class StructuredContextRequirement:
    """Source-cell context needed to resolve one structured reference."""

    uses_current_table: bool = False
    uses_current_table_row: bool = False
    current_row_tables: tuple[TableBinding, ...] = ()


@dataclass(frozen=True, slots=True)
class ReferenceContext:
    """Immutable workbook catalog used by formula analysis."""

    sheets: tuple[SheetDescriptor, ...]
    defined_names: tuple[DefinedName, ...] = ()
    tables: tuple[TableBinding, ...] = ()
    external_links: Mapping[int, str] = field(default_factory=_empty_external_links)
    _sheets_by_name: Mapping[str, SheetDescriptor] = field(init=False, repr=False)
    _names_by_key: Mapping[NameKey, DefinedName] = field(init=False, repr=False)
    _tables_by_name: Mapping[str, TableBinding] = field(init=False, repr=False)
    _table_rects: Mapping[TableBinding, Rect] = field(init=False, repr=False)
    _table_indexes: Mapping[int, _TableSpatialNode | None] = field(init=False, repr=False)

    def __post_init__(self) -> None:
        sheet_map: dict[str, SheetDescriptor] = {}
        for sheet in self.sheets:
            key = sheet.name.casefold()
            if key in sheet_map:
                raise ValueError(f"duplicate sheet name in reference context: {sheet.name}")
            sheet_map[key] = sheet

        name_map: dict[NameKey, DefinedName] = {}
        for defined_name in self.defined_names:
            key = (defined_name.scope_sheet_order, defined_name.name.casefold())
            if key in name_map:
                raise ValueError(f"duplicate defined name in one scope: {defined_name.name}")
            name_map[key] = defined_name

        table_map: dict[str, TableBinding] = {}
        table_rects: dict[TableBinding, Rect] = {}
        table_entries_by_sheet: dict[int, list[tuple[TableBinding, Rect]]] = {}
        for binding in self.tables:
            sheet = sheet_map.get(binding.sheet_name.casefold())
            if sheet is None or sheet.order != binding.sheet_order:
                raise ValueError(f"table uses an unknown sheet: {binding.sheet_name}")
            for alias in {binding.table.name.casefold(), binding.table.display_name.casefold()}:
                previous = table_map.get(alias)
                if previous is not None and previous != binding:
                    raise ValueError(f"duplicate ListObject name: {binding.table.name}")
                table_map[alias] = binding
            rect = parse_rect(binding.table.ref)
            table_rects[binding] = rect
            table_entries_by_sheet.setdefault(binding.sheet_order, []).append((binding, rect))

        links: dict[int, str] = {}
        for index, target in self.external_links.items():
            if type(index) is not int or index < 1 or not target:
                raise ValueError("external links require positive integer ids and targets")
            links[index] = target
        object.__setattr__(self, "external_links", MappingProxyType(links))
        object.__setattr__(self, "_sheets_by_name", MappingProxyType(sheet_map))
        object.__setattr__(self, "_names_by_key", MappingProxyType(name_map))
        object.__setattr__(self, "_tables_by_name", MappingProxyType(table_map))
        object.__setattr__(self, "_table_rects", MappingProxyType(table_rects))
        object.__setattr__(
            self,
            "_table_indexes",
            MappingProxyType(
                {
                    sheet_order: _build_table_spatial_tree(tuple(entries))
                    for sheet_order, entries in table_entries_by_sheet.items()
                }
            ),
        )

    def sheet_named(self, name: str) -> SheetDescriptor | None:
        """Resolve a sheet case-insensitively."""
        return self._sheets_by_name.get(name.casefold())

    def defined_name(self, name: str, sheet_order: int) -> DefinedName | None:
        """Resolve sheet scope before workbook scope."""
        folded = name.casefold()
        return self._names_by_key.get((sheet_order, folded)) or self._names_by_key.get(
            (None, folded)
        )

    def table_named(self, name: str) -> TableBinding | None:
        """Resolve a ListObject name or display name case-insensitively."""
        return self._tables_by_name.get(name.casefold())

    def current_table(self, anchor: FormulaAnchor) -> TableBinding | None:
        """Return the unique table containing the formula cell, if any."""
        matches = self.tables_intersecting(
            anchor.sheet_order,
            Rect(anchor.row, anchor.row, anchor.col, anchor.col),
        )
        return matches[0] if len(matches) == 1 else None

    def table_rect(self, binding: TableBinding) -> Rect:
        """Return one validated ListObject rectangle."""
        try:
            return self._table_rects[binding]
        except KeyError as exc:
            raise ValueError("table binding is not part of this reference context") from exc

    def tables_intersecting(
        self,
        sheet_order: int,
        rect: Rect,
    ) -> tuple[TableBinding, ...]:
        """Return deterministic exact ListObject intersections on one sheet."""
        root = self._table_indexes.get(sheet_order)
        if root is None:
            return ()
        result: list[TableBinding] = []
        pending = [root]
        while pending:
            node = pending.pop()
            if not node.bounds.intersects(rect):
                continue
            for binding, table_rect in node.entries:
                if table_rect.intersects(rect):
                    result.append(binding)
            if node.right is not None:
                pending.append(node.right)
            if node.left is not None:
                pending.append(node.left)
        return tuple(
            sorted(
                result,
                key=lambda binding: (
                    self._table_rects[binding].row_min,
                    self._table_rects[binding].col_min,
                    self._table_rects[binding].row_max,
                    self._table_rects[binding].col_max,
                    binding.table.name.casefold(),
                ),
            )
        )


@contextmanager
def name_expansion_scope() -> Generator[None, None, None]:
    try:
        task = current_task()
    except RuntimeError:
        task = None
    owner = (current_thread(), task)
    existing = _NAME_EXPANSION_CACHE.get()
    if existing is not None and existing.owner == owner:
        yield
        return
    scope = _NAME_EXPANSION_CACHE.set(_NameExpansionState(owner))
    try:
        yield
    finally:
        _NAME_EXPANSION_CACHE.reset(scope)


def classify_ref(
    text: str,
    *,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    function_position: bool = False,
    name_stack: tuple[NameKey, ...] = (),
) -> ReferenceClassification:
    with name_expansion_scope():
        return _classify_ref_impl(
            text,
            anchor=anchor,
            context=context,
            function_position=function_position,
            name_stack=name_stack,
        )


def _classify_ref_impl(
    text: str,
    *,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    function_position: bool = False,
    name_stack: tuple[NameKey, ...] = (),
) -> ReferenceClassification:
    """Classify one tokenizer RANGE operand or non-built-in callable name."""
    original = text
    value = text.strip()
    if not value:
        return _parse_problem(original, "Formula contains an empty reference operand.")

    modern_range = parse_modern_a1_range(value, anchor.cell)
    if modern_range is not None:
        parsed = modern_range.parsed
        if not (modern_range.left_spill or modern_range.right_spill):
            return _classify_a1(
                original,
                parsed.qualifier,
                parsed.geometry,
                False,
                anchor,
                context,
            )
        left_geometry, right_geometry = modern_range_endpoint_geometries(modern_range)
        return _merge_classifications(
            (
                _classify_a1(
                    original,
                    parsed.qualifier,
                    left_geometry,
                    modern_range.left_spill,
                    anchor,
                    context,
                ),
                _classify_a1(
                    original,
                    parsed.qualifier,
                    right_geometry,
                    modern_range.right_spill,
                    anchor,
                    context,
                ),
            )
        )

    reference_range = _reference_range_for_context(value, anchor, context)
    if reference_range is not None:
        return _classify_reference_range(
            original,
            reference_range,
            anchor,
            context,
            name_stack,
        )

    implicit_intersection = value.startswith("@")
    if implicit_intersection:
        value = value[1:]
    spill = value.endswith("#")
    if spill:
        value = value[:-1]
    if not value:
        return _parse_problem(original, "Formula contains an incomplete reference operator.")

    parsed = parse_a1_reference(value, anchor.cell)
    if parsed is not None:
        return _classify_a1(original, parsed.qualifier, parsed.geometry, spill, anchor, context)

    if "[" in value or value.startswith("@["):
        return _classify_structured(original, value, spill, anchor, context)

    if _R1C1_REFERENCE.fullmatch(value):
        return ReferenceClassification(
            references=(ExtractedReference(original, None, None, None, "opaque:ref"),),
            opaque=True,
        )

    return _classify_name(
        original,
        value,
        spill,
        anchor,
        context,
        function_position=function_position,
        name_stack=name_stack,
    )


def structured_context_requirement(
    text: str,
    context: ReferenceContext,
) -> StructuredContextRequirement:
    """Describe whether a structured expression changes meaning across source cells."""
    value = text.strip()
    if value.startswith("@"):
        value = value[1:]
    if value.endswith("#"):
        value = value[:-1]
    # Parenthesized unions and whitespace intersections are tokenized into
    # their component RANGE operands. A plain colon composite is retained as
    # one RANGE token, so split that form through the same quote/bracket/3-D
    # aware candidate parser used by classification.
    try:
        tokens = tokenize_formula(value)
    except FormulaSyntaxError:
        tokens = ()
    operands = tuple(
        token.value for token in tokens if token.type == "OPERAND" and token.subtype == "RANGE"
    )
    if operands and operands != (value,):
        return _merge_structured_context_requirements(
            structured_context_requirement(operand, context) for operand in operands
        )

    candidates = parse_reference_range_candidates(value, CellRef(1, 1))
    endpoints = tuple(
        dict.fromkeys(
            endpoint.original
            for candidate in candidates
            for endpoint in (candidate.left, candidate.right)
        )
    )
    if endpoints:
        return _merge_structured_context_requirements(
            structured_context_requirement(endpoint, context) for endpoint in endpoints
        )

    split = _split_structured(value)
    if split is None:
        return StructuredContextRequirement()
    table_name, selectors = split
    uses_current_row = any(
        not selector.escaped_first and selector.text.startswith("@") for selector in selectors
    )
    if table_name is None:
        return StructuredContextRequirement(
            uses_current_table=True,
            uses_current_table_row=uses_current_row,
        )
    if uses_current_row:
        binding = context.table_named(table_name)
        return StructuredContextRequirement(
            current_row_tables=() if binding is None else (binding,),
        )
    return StructuredContextRequirement()


def _merge_structured_context_requirements(
    requirements: Iterable[StructuredContextRequirement],
) -> StructuredContextRequirement:
    materialized = tuple(requirements)
    return StructuredContextRequirement(
        uses_current_table=any(item.uses_current_table for item in materialized),
        uses_current_table_row=any(item.uses_current_table_row for item in materialized),
        current_row_tables=tuple(
            dict.fromkeys(table for item in materialized for table in item.current_row_tables)
        ),
    )


def _reference_range_for_context(
    text: str,
    anchor: FormulaAnchor,
    context: ReferenceContext,
) -> ParsedReferenceRange | None:
    candidates = parse_reference_range_candidates(text, anchor.cell)
    if not candidates:
        return None
    scores = tuple(
        sum(
            endpoint.parsed_a1 is None
            and context.defined_name(endpoint.core, anchor.sheet_order) is not None
            for endpoint in (candidate.left, candidate.right)
        )
        for candidate in candidates
    )
    best = max(scores)
    if best:
        winners = [
            candidate for candidate, score in zip(candidates, scores, strict=True) if score == best
        ]
        lexical = parse_reference_range(text, anchor.cell)
        if lexical in winners:
            return lexical
        return winners[0]
    return parse_reference_range(text, anchor.cell)


def _classify_reference_range(
    token: str,
    parsed: ParsedReferenceRange,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    name_stack: tuple[NameKey, ...],
) -> ReferenceClassification:
    endpoints = (parsed.left, parsed.right)
    classifications = tuple(
        classify_ref(
            endpoint.original,
            anchor=anchor,
            context=context,
            name_stack=name_stack,
        )
        for endpoint in endpoints
    )
    merged = _retoken_classification(_merge_classifications(classifications), token)
    computed_endpoint = any(classification.non_range_name for classification in classifications)
    compatible = not computed_endpoint and _single_area_endpoints_are_compatible(*classifications)

    if compatible and not any(endpoint.spill for endpoint in endpoints):
        component_geometries = tuple(
            reference.geometry
            for classification in classifications
            for reference in classification.references
            if reference.geometry is not None
        )
        assert len(component_geometries) == 2
        a1_indexes = [
            index for index, endpoint in enumerate(endpoints) if endpoint.parsed_a1 is not None
        ]
        if len(a1_indexes) == 1:
            a1_index = a1_indexes[0]
            name_index = 1 - a1_index
            a1_reference = classifications[a1_index].references[0]
            name_reference = classifications[name_index].references[0]
            a1_endpoint = endpoints[a1_index]
            assert a1_endpoint.parsed_a1 is not None
            geometry = _bounding_name_a1_geometry(
                name_reference.geometry,
                a1_endpoint.parsed_a1.geometry,
                anchor.cell,
            )
            if a1_reference.via == "ref" and geometry is not None:
                synthesized = ExtractedReference(
                    token,
                    name_reference.dst_sheet_order,
                    name_reference.dst_sheet_name,
                    geometry,
                    name_reference.via,
                    component_geometries,
                )
                return replace(merged, references=(synthesized,))
        elif not a1_indexes:
            geometry = _bounding_absolute_geometries(
                classifications[0].references[0].geometry,
                classifications[1].references[0].geometry,
            )
            if geometry is not None:
                references = tuple(
                    ExtractedReference(
                        token,
                        reference.dst_sheet_order,
                        reference.dst_sheet_name,
                        geometry,
                        reference.via,
                        component_geometries,
                    )
                    for reference in (
                        classifications[0].references[0],
                        classifications[1].references[0],
                    )
                )
                return replace(
                    merged,
                    references=tuple(dict.fromkeys(references)),
                )
        geometry = _bounding_geometries_at_anchor(
            classifications[0].references[0].geometry,
            classifications[1].references[0].geometry,
            anchor.cell,
        )
        if geometry is not None:
            endpoint_references = (
                classifications[0].references[0],
                classifications[1].references[0],
            )
            non_plain = tuple(
                reference for reference in endpoint_references if reference.via != "ref"
            )
            retained = non_plain or endpoint_references[:1]
            references = tuple(
                ExtractedReference(
                    token,
                    reference.dst_sheet_order,
                    reference.dst_sheet_name,
                    geometry,
                    reference.via,
                    component_geometries,
                )
                for reference in retained
            )
            return replace(
                merged,
                references=tuple(dict.fromkeys(references)),
            )

    if computed_endpoint:
        references = list(merged.references)
        issues = list(merged.issues)
        for label in merged.dynamic_result_labels:
            marker = ExtractedReference(token, None, None, None, f"opaque:{label}")
            if marker not in references and not any(
                reference.via == marker.via for reference in references
            ):
                references.append(marker)
            issue = _issue(
                "info",
                "I_DYNAMIC_REF",
                f"Function {label} returns a dynamic reference.",
                function=label,
            )
            if not any(
                existing.code == issue.code and existing.related.get("function") == label
                for existing in issues
            ):
                issues.append(issue)
        if not any(reference.via.startswith("opaque:") for reference in references):
            references.append(ExtractedReference(token, None, None, None, "opaque:ref"))
        return replace(
            merged,
            references=tuple(references),
            issues=tuple(issues),
            opaque=True,
        )

    # A multi-area name or cross-sheet/3-D/external endpoint cannot be encoded
    # as one exact rectangle. Keep its endpoint edges, but advertise that the
    # result is conservative rather than exact range geometry.
    if not compatible or not any(endpoint.spill for endpoint in endpoints):
        return replace(merged, opaque=True)
    return merged


def _single_area_endpoints_are_compatible(
    left: ReferenceClassification,
    right: ReferenceClassification,
) -> bool:
    if len(left.references) != 1 or len(right.references) != 1:
        return False
    left_reference = left.references[0]
    right_reference = right.references[0]
    return bool(
        left_reference.geometry is not None
        and right_reference.geometry is not None
        and left_reference.dst_sheet_order is not None
        and left_reference.dst_sheet_order == right_reference.dst_sheet_order
    )


def _bounding_absolute_geometries(
    left: ReferenceGeometry | None,
    right: ReferenceGeometry | None,
) -> ReferenceGeometry | None:
    if left is None or right is None or _geometry_is_relative(left) or _geometry_is_relative(right):
        return None
    return ReferenceGeometry(
        AxisTerm(
            False, min(left.row_a.value, left.row_b.value, right.row_a.value, right.row_b.value)
        ),
        AxisTerm(
            False, max(left.row_a.value, left.row_b.value, right.row_a.value, right.row_b.value)
        ),
        AxisTerm(
            False, min(left.col_a.value, left.col_b.value, right.col_a.value, right.col_b.value)
        ),
        AxisTerm(
            False, max(left.col_a.value, left.col_b.value, right.col_a.value, right.col_b.value)
        ),
    )


def _bounding_geometries_at_anchor(
    left: ReferenceGeometry | None,
    right: ReferenceGeometry | None,
    anchor: CellRef,
) -> ReferenceGeometry | None:
    if left is None or right is None:
        return None
    rows = _bounding_axis_at_anchor(
        (left.row_a, left.row_b, right.row_a, right.row_b),
        anchor.row,
    )
    columns = _bounding_axis_at_anchor(
        (left.col_a, left.col_b, right.col_a, right.col_b),
        anchor.col,
    )
    return ReferenceGeometry(rows[0], rows[1], columns[0], columns[1])


def _bounding_axis_at_anchor(
    terms: tuple[AxisTerm, AxisTerm, AxisTerm, AxisTerm],
    anchor: int,
) -> tuple[AxisTerm, AxisTerm]:
    def resolved(term: AxisTerm) -> int:
        return anchor + term.value if term.relative else term.value

    lower = min(terms, key=lambda term: (resolved(term), not term.relative))
    upper = max(terms, key=lambda term: (resolved(term), term.relative))
    return lower, upper


def _bounding_name_a1_geometry(
    name_geometry: ReferenceGeometry | None,
    a1_geometry: ReferenceGeometry,
    anchor: CellRef,
) -> ReferenceGeometry | None:
    if name_geometry is None:
        return None
    rows = _bounding_name_axis(
        name_geometry.row_a,
        name_geometry.row_b,
        a1_geometry.row_a,
        a1_geometry.row_b,
        anchor.row,
    )
    columns = _bounding_name_axis(
        name_geometry.col_a,
        name_geometry.col_b,
        a1_geometry.col_a,
        a1_geometry.col_b,
        anchor.col,
    )
    if rows is None or columns is None:
        return None
    return ReferenceGeometry(rows[0], rows[1], columns[0], columns[1])


def _bounding_name_axis(
    name_a: AxisTerm,
    name_b: AxisTerm,
    point_a: AxisTerm,
    point_b: AxisTerm,
    anchor: int,
) -> tuple[AxisTerm, AxisTerm] | None:
    if name_a == name_b and point_a == point_b:
        if name_a.relative == point_a.relative:
            if name_a.value <= point_a.value:
                return name_a, point_a
            return point_a, name_a
        return point_a, name_a
    if name_a.relative or name_b.relative:
        return None
    name_min, name_max = sorted((name_a.value, name_b.value))
    if not point_a.relative and not point_b.relative:
        return (
            AxisTerm(False, min(name_min, point_a.value, point_b.value)),
            AxisTerm(False, max(name_max, point_a.value, point_b.value)),
        )
    if point_a != point_b:
        point_a_value = anchor + point_a.value if point_a.relative else point_a.value
        point_b_value = anchor + point_b.value if point_b.relative else point_b.value
        lower = point_a if point_a_value <= point_b_value else point_b
        upper = point_b if point_a_value <= point_b_value else point_a
        if min(point_a_value, point_b_value) >= name_min:
            lower = AxisTerm(False, name_min)
        if max(point_a_value, point_b_value) <= name_max:
            upper = AxisTerm(False, name_max)
        return lower, upper
    point = anchor + point_a.value if point_a.relative else point_a.value
    if point <= name_min:
        return point_a, AxisTerm(False, name_max)
    if point >= name_max:
        return AxisTerm(False, name_min), point_a
    # A point strictly inside the name contributes no additional bounds.
    return AxisTerm(False, name_min), AxisTerm(False, name_max)


def _retoken_classification(
    classification: ReferenceClassification,
    token: str,
) -> ReferenceClassification:
    references: list[ExtractedReference] = []
    for reference in classification.references:
        retokened = replace(reference, token=token)
        if retokened not in references:
            references.append(retokened)
    return replace(classification, references=tuple(references))


def _classify_a1(
    token: str,
    qualifier: str,
    geometry: ReferenceGeometry,
    spill: bool,
    anchor: FormulaAnchor,
    context: ReferenceContext,
) -> ReferenceClassification:
    via = "spill" if spill else "ref"
    if not qualifier:
        return ReferenceClassification(
            references=(
                ExtractedReference(token, anchor.sheet_order, anchor.sheet_name, geometry, via),
            )
        )

    raw_qualifier = qualifier[:-1]
    sheet_expression = _unquote_sheet_expression(raw_qualifier)
    if sheet_expression is None:
        return _parse_problem(token, "Formula contains a malformed quoted sheet name.")

    external = _EXTERNAL_BOOK.search(sheet_expression)
    if external is not None:
        label = external.group(1)
        if label.isdecimal():
            target = context.external_links.get(int(label), f"[{label}]")
        else:
            target = sheet_expression[: external.end()]
        return ReferenceClassification(
            references=(
                ExtractedReference(
                    token,
                    None,
                    None,
                    None,
                    f"external:{external_link_label(target)}",
                ),
            )
        )

    if ":" in sheet_expression:
        first_name, second_name = sheet_expression.split(":", 1)
        first = context.sheet_named(first_name)
        second = context.sheet_named(second_name)
        if first is None or second is None:
            return _parse_problem(token, "A 3-D reference names a worksheet that does not exist.")
        low, high = sorted((first.order, second.order))
        sheets = sorted(
            (sheet for sheet in context.sheets if low <= sheet.order <= high),
            key=lambda sheet: sheet.order,
        )
        return ReferenceClassification(
            references=tuple(
                ExtractedReference(
                    token,
                    sheet.order,
                    sheet.name,
                    geometry,
                    "spill" if spill else "3d",
                )
                for sheet in sheets
            )
        )

    sheet = context.sheet_named(sheet_expression)
    if sheet is None:
        return _parse_problem(token, f"Formula names an unknown worksheet: {sheet_expression}")
    return ReferenceClassification(
        references=(ExtractedReference(token, sheet.order, sheet.name, geometry, via),)
    )


def _classify_structured(
    token: str,
    value: str,
    spill: bool,
    anchor: FormulaAnchor,
    context: ReferenceContext,
) -> ReferenceClassification:
    split = _split_structured(value)
    if split is None:
        return _structured_problem(token, "Unsupported structured-reference syntax.")
    table_name, selectors = split
    binding = (
        context.current_table(anchor) if table_name is None else context.table_named(table_name)
    )
    if binding is None:
        return _structured_problem(token, "Structured reference does not identify a table.")

    table_rect = parse_rect(binding.table.ref)
    row_selector: str | None = None
    column_name: str | None = None
    for selector_info in selectors:
        selector = selector_info.text
        if not selector_info.escaped_first and selector.startswith("@"):
            if row_selector is not None:
                return _structured_problem(
                    token, "Structured reference has conflicting row selectors."
                )
            row_selector = "@"
            if len(selector) > 1:
                column_name = selector[1:]
        elif not selector_info.escaped_first and selector.casefold() in {
            "#all",
            "#data",
            "#headers",
            "#totals",
        }:
            if row_selector is not None:
                return _structured_problem(
                    token, "Structured reference has conflicting row selectors."
                )
            row_selector = selector.casefold()
        elif not selector_info.escaped_first and selector.startswith("#"):
            return _structured_problem(token, f"Unsupported structured selector: {selector}")
        elif column_name is None:
            column_name = selector
        else:
            return _structured_problem(token, "Structured column ranges are not supported in v0.1.")

    column_bounds = (table_rect.col_min, table_rect.col_max)
    if column_name is not None:
        try:
            column_offset = next(
                index
                for index, name in enumerate(binding.table.columns)
                if name.casefold() == column_name.casefold()
            )
        except StopIteration:
            return _structured_problem(token, f"Unknown table column: {column_name}")
        column = table_rect.col_min + column_offset
        column_bounds = (column, column)

    data_min = table_rect.row_min + binding.table.header_rows
    data_max = table_rect.row_max - binding.table.totals_rows
    selector = row_selector or "#data"
    if selector == "@":
        if binding.sheet_order != anchor.sheet_order or not data_min <= anchor.row <= data_max:
            return _structured_problem(
                token, "Current-row reference is outside the table data body."
            )
        row_a = row_b = AxisTerm(True, 0)
    elif selector == "#all":
        row_a = AxisTerm(False, table_rect.row_min)
        row_b = AxisTerm(False, table_rect.row_max)
    elif selector == "#data":
        if data_min > data_max:
            return _structured_problem(token, "Table has no data rows.")
        row_a = AxisTerm(False, data_min)
        row_b = AxisTerm(False, data_max)
    elif selector == "#headers":
        if binding.table.header_rows < 1:
            return _structured_problem(token, "Table has no header row.")
        row_a = AxisTerm(False, table_rect.row_min)
        row_b = AxisTerm(False, data_min - 1)
    else:
        if binding.table.totals_rows < 1:
            return _structured_problem(token, "Table has no totals row.")
        row_a = AxisTerm(False, data_max + 1)
        row_b = AxisTerm(False, table_rect.row_max)

    canonical_column = "" if column_name is None else f"[{column_name}]"
    reference = ExtractedReference(
        token,
        binding.sheet_order,
        binding.sheet_name,
        ReferenceGeometry(
            row_a,
            row_b,
            AxisTerm(False, column_bounds[0]),
            AxisTerm(False, column_bounds[1]),
        ),
        "spill" if spill else f"structured:{binding.table.name}{canonical_column}",
    )
    return ReferenceClassification(references=(reference,))


def _classify_name(
    token: str,
    name: str,
    spill: bool,
    anchor: FormulaAnchor,
    context: ReferenceContext,
    *,
    function_position: bool,
    name_stack: tuple[NameKey, ...],
) -> ReferenceClassification:
    defined_name = context.defined_name(name, anchor.sheet_order)
    if defined_name is None:
        position = "function" if function_position else "operand"
        issue = _issue(
            "warn",
            "W_UNKNOWN_NAME",
            f"Unknown defined name in {position} position: {name}",
            name=name,
        )
        return ReferenceClassification(
            references=(ExtractedReference(token, None, None, None, "opaque:name"),),
            issues=(issue,),
            opaque=True,
        )

    key = (defined_name.scope_sheet_order, defined_name.name.casefold())
    state = _NAME_EXPANSION_CACHE.get()
    cache = state.cache if state is not None else None
    cache_key: _NameCacheKey = (
        _IdentityKey(context),
        key,
        anchor.sheet_order,
        anchor.sheet_name,
        anchor.row,
        anchor.col,
        token,
        spill,
        function_position,
        name_stack,
    )
    if cache is not None and cache_key in cache:
        return cache[cache_key]

    def remember(result: ReferenceClassification) -> ReferenceClassification:
        if cache is not None:
            cache[cache_key] = result
        return result

    via = "spill" if spill else f"name:{defined_name.name}"
    if defined_name.kind in {"range", "multi_range"}:
        references: list[ExtractedReference] = []
        for area in defined_name.areas:
            sheet = context.sheet_named(area.sheet_name)
            if sheet is None:
                return remember(
                    _parse_problem(token, "Defined-name area uses an unknown worksheet.")
                )
            references.append(
                ExtractedReference(
                    token,
                    sheet.order,
                    sheet.name,
                    _absolute_geometry(area.rect),
                    via,
                )
            )
        return remember(ReferenceClassification(references=tuple(references)))
    if defined_name.kind == "constant":
        return remember(ReferenceClassification(non_range_name=True))

    if key in name_stack or len(name_stack) >= 32:
        issue = _issue(
            "warn",
            "W_PARSE",
            f"Defined-name expansion is recursive: {defined_name.name}",
            name=defined_name.name,
        )
        return remember(
            ReferenceClassification(
                references=(ExtractedReference(token, None, None, None, "opaque:name"),),
                issues=(issue,),
                opaque=True,
                non_range_name=True,
            )
        )

    # Local import avoids a module cycle while keeping all formula/name parsing
    # behind the same failure-containment and LET/LAMBDA scope machinery.
    from excel_lsp.core.formulas.analysis import analyze_formula

    analysis = analyze_formula(
        defined_name.refers_to,
        anchor=anchor,
        context=context,
        name_stack=(*name_stack, key),
    )
    expanded: list[ExtractedReference] = []
    relative_seen = False
    for reference in analysis.references:
        relative_geometry = (
            reference.geometry is not None and _geometry_is_relative(reference.geometry)
        ) or any(_geometry_is_relative(geometry) for geometry in reference.extrusion_geometries)
        if relative_geometry:
            expanded.append(ExtractedReference(reference.token, None, None, None, "name-relative"))
            relative_seen = True
        else:
            expanded.append(reference)
    return remember(
        ReferenceClassification(
            references=tuple(expanded),
            issues=analysis.issues,
            function_calls=analysis.function_calls,
            volatile=analysis.volatile,
            opaque=analysis.opaque or relative_seen,
            non_range_name=True,
            computed_returns_reference=analysis.returns_reference,
            dynamic_result_labels=analysis.dynamic_result_labels,
        )
    )


def _split_structured(
    value: str,
) -> tuple[str | None, tuple[_StructuredSelector, ...]] | None:
    opening = value.find("[")
    if opening < 0 or not value.endswith("]"):
        return None
    table_name = value[:opening] or None
    specification = value[opening:]
    inner = specification[1:-1]
    if not inner:
        return None
    if inner.startswith("@["):
        parsed = _consume_structured_selector(inner, 1)
        if parsed is None or parsed[1] != len(inner):
            return None
        selector, _position = parsed
        return table_name, (_StructuredSelector(f"@{selector.text}"),)
    if not inner.startswith("["):
        decoded = _decode_structured_selector(inner)
        if decoded is None:
            return None
        decoded_text, escaped_first = decoded
        return table_name, (_StructuredSelector(decoded_text, escaped_first),)

    selectors: list[_StructuredSelector] = []
    position = 0
    while position < len(inner):
        parsed = _consume_structured_selector(inner, position)
        if parsed is None:
            return None
        selector, position = parsed
        selectors.append(selector)
        if position == len(inner):
            break
        if inner[position] != ",":
            return None
        position += 1
    return table_name, tuple(selectors)


def _consume_structured_selector(
    value: str,
    position: int,
) -> tuple[_StructuredSelector, int] | None:
    if position >= len(value) or value[position] != "[":
        return None
    characters: list[str] = []
    escaped_first = False
    position += 1
    while position < len(value):
        character = value[position]
        if character == "'":
            if position + 1 < len(value) and value[position + 1] in "[]#'@":
                if not characters:
                    escaped_first = True
                characters.append(value[position + 1])
                position += 2
                continue
            return None
        if character == "]":
            if not characters:
                return None
            return _StructuredSelector("".join(characters), escaped_first), position + 1
        if character == "[":
            return None
        if character in "#@" and characters:
            return None
        characters.append(character)
        position += 1
    return None


def _decode_structured_selector(value: str) -> tuple[str, bool] | None:
    characters: list[str] = []
    escaped_first = False
    position = 0
    while position < len(value):
        character = value[position]
        if character == "'":
            if position + 1 < len(value) and value[position + 1] in "[]#'@":
                if not characters:
                    escaped_first = True
                characters.append(value[position + 1])
                position += 2
                continue
            return None
        if character in "[]" or (character in "#@" and characters):
            return None
        characters.append(character)
        position += 1
    return "".join(characters), escaped_first


def _build_table_spatial_tree(
    entries: tuple[tuple[TableBinding, Rect], ...],
) -> _TableSpatialNode | None:
    if not entries:
        return None
    bounds = _table_bounds(entries)
    if len(entries) <= 8:
        return _TableSpatialNode(bounds, entries=entries)
    row_centers = tuple(rect.row_min + rect.row_max for _binding, rect in entries)
    col_centers = tuple(rect.col_min + rect.col_max for _binding, rect in entries)
    if max(row_centers) - min(row_centers) >= max(col_centers) - min(col_centers):
        ordered = sorted(
            entries,
            key=lambda item: (
                item[1].row_min + item[1].row_max,
                item[1].col_min + item[1].col_max,
                item[1].row_min,
                item[1].col_min,
                item[0].table.name.casefold(),
            ),
        )
    else:
        ordered = sorted(
            entries,
            key=lambda item: (
                item[1].col_min + item[1].col_max,
                item[1].row_min + item[1].row_max,
                item[1].row_min,
                item[1].col_min,
                item[0].table.name.casefold(),
            ),
        )
    middle = len(ordered) // 2
    return _TableSpatialNode(
        bounds,
        left=_build_table_spatial_tree(tuple(ordered[:middle])),
        right=_build_table_spatial_tree(tuple(ordered[middle:])),
    )


def _table_bounds(entries: tuple[tuple[TableBinding, Rect], ...]) -> Rect:
    return Rect(
        min(rect.row_min for _binding, rect in entries),
        max(rect.row_max for _binding, rect in entries),
        min(rect.col_min for _binding, rect in entries),
        max(rect.col_max for _binding, rect in entries),
    )


def _unquote_sheet_expression(value: str) -> str | None:
    if not value.startswith("'"):
        return value if "'" not in value else None
    if not value.endswith("'"):
        return None
    return value[1:-1].replace("''", "'")


def _absolute_geometry(rect: Rect) -> ReferenceGeometry:
    return ReferenceGeometry(
        AxisTerm(False, rect.row_min),
        AxisTerm(False, rect.row_max),
        AxisTerm(False, rect.col_min),
        AxisTerm(False, rect.col_max),
    )


def _geometry_is_relative(geometry: ReferenceGeometry) -> bool:
    return any(
        term.relative for term in (geometry.row_a, geometry.row_b, geometry.col_a, geometry.col_b)
    )


def _issue(
    severity: IssueSeverity,
    code: str,
    message: str,
    **related: object,
) -> FormulaIssue:
    return FormulaIssue(severity, code, message, related)


def _merge_classifications(
    classifications: tuple[ReferenceClassification, ...],
) -> ReferenceClassification:
    references: list[ExtractedReference] = []
    issues: list[FormulaIssue] = []
    function_calls: list[str] = []
    for classification in classifications:
        for reference in classification.references:
            if reference not in references:
                references.append(reference)
        for issue in classification.issues:
            if issue not in issues:
                issues.append(issue)
        function_calls.extend(classification.function_calls)
    return ReferenceClassification(
        references=tuple(references),
        issues=tuple(issues),
        function_calls=tuple(function_calls),
        volatile=any(classification.volatile for classification in classifications),
        opaque=any(classification.opaque for classification in classifications),
        non_range_name=any(classification.non_range_name for classification in classifications),
        computed_returns_reference=any(
            classification.computed_returns_reference for classification in classifications
        ),
        dynamic_result_labels=tuple(
            dict.fromkeys(
                label
                for classification in classifications
                for label in classification.dynamic_result_labels
            )
        ),
    )


def _parse_problem(token: str, message: str) -> ReferenceClassification:
    return ReferenceClassification(
        references=(ExtractedReference(token, None, None, None, "opaque:parse"),),
        issues=(_issue("warn", "W_PARSE", message, token=token),),
        opaque=True,
    )


def _structured_problem(token: str, message: str) -> ReferenceClassification:
    result = _parse_problem(token, message)
    return replace(
        result,
        references=(ExtractedReference(token, None, None, None, "opaque:structured"),),
    )


__all__ = [
    "ExtractedReference",
    "FormulaAnchor",
    "FormulaIssue",
    "ReferenceClassification",
    "ReferenceContext",
    "StructuredContextRequirement",
    "TableBinding",
    "classify_ref",
    "structured_context_requirement",
]
