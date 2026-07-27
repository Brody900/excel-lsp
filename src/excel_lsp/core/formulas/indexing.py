"""Workbook-sheet formula analysis independent of SQLite persistence."""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field, replace
from types import MappingProxyType
from typing import Literal, TypeVar

from excel_lsp.core.formulas.a1 import extrude_reference, resolve_reference
from excel_lsp.core.formulas.analysis import ReferenceUse, analyze_formula
from excel_lsp.core.formulas.blocks import (
    FormulaBlock,
    FormulaCell,
    build_formula_blocks,
    detect_inconsistent_formulas,
    normalize_formula_cells,
)
from excel_lsp.core.formulas.references import (
    ExtractedReference,
    FormulaAnchor,
    FormulaIssue,
    ReferenceContext,
    TableBinding,
    structured_context_requirement,
)
from excel_lsp.core.formulas.translation import translate_a1_formula
from excel_lsp.core.models import Rect, SheetDescriptor
from excel_lsp.core.parse.coordinates import make_cell_ref
from excel_lsp.core.symbols import cell_symbol_id, formula_block_symbol_id

DiagnosticSeverity = Literal["info", "warn"]
_P3_ISSUE_CODES = frozenset({"I_DYNAMIC_REF", "W_PARSE", "W_UNKNOWN_NAME"})
_T = TypeVar("_T")


def _empty_related() -> Mapping[str, object]:
    return MappingProxyType({})


@dataclass(frozen=True, slots=True)
class FormulaEdge:
    """One formula-block dependency ready for schema persistence."""

    source_block_n: int
    dst_sheet_order: int | None
    rect: Rect | None
    via: str

    def __post_init__(self) -> None:
        if type(self.source_block_n) is not int or self.source_block_n < 0:
            raise ValueError("formula edge source ordinal must be nonnegative")
        if (self.dst_sheet_order is None) != (self.rect is None):
            raise ValueError("formula edge destination sheet and rectangle must coexist")
        if not self.via:
            raise ValueError("formula edge via must not be empty")


@dataclass(frozen=True, slots=True)
class FormulaDiagnostic:
    """One P3-owned diagnostic ready for schema persistence."""

    severity: DiagnosticSeverity
    code: str
    row: int
    col: int
    ref: str
    message: str
    related: Mapping[str, object] = field(default_factory=_empty_related)

    def __post_init__(self) -> None:
        if self.severity not in {"info", "warn"}:
            raise ValueError("P3 diagnostics must be info or warn")
        if not self.code or not self.ref or not self.message:
            raise ValueError("formula diagnostic text fields must not be empty")
        object.__setattr__(self, "related", MappingProxyType(dict(self.related)))


@dataclass(frozen=True, slots=True)
class SheetFormulaAnalysis:
    """Complete deterministic P3 result for one worksheet."""

    blocks: tuple[FormulaBlock, ...]
    edges: tuple[FormulaEdge, ...]
    diagnostics: tuple[FormulaDiagnostic, ...]


@dataclass(frozen=True, slots=True)
class _SemanticTile:
    rect: Rect
    context_key: tuple[str, ...]


def analyze_sheet_formulas(
    descriptor: SheetDescriptor,
    formula_cells: Iterable[FormulaCell],
    regions: Sequence[Rect],
    context: ReferenceContext,
) -> SheetFormulaAnalysis:
    """Build exact blocks, extract one reference set per block, and diagnose."""
    patterns = normalize_formula_cells(formula_cells)
    base_blocks = build_formula_blocks(patterns)
    blocks: list[FormulaBlock] = []
    edges: list[FormulaEdge] = []
    diagnostics: list[FormulaDiagnostic] = []

    for block in base_blocks:
        block_anchor = FormulaAnchor(
            descriptor.order,
            descriptor.name,
            block.rect.row_min,
            block.rect.col_min,
        )
        initial_semantic = analyze_formula(
            block.anchor_formula,
            anchor=block_anchor,
            context=context,
        )
        contextual_uses = tuple(
            use
            for use in initial_semantic.reference_uses
            if _reference_use_requires_structured_context(use, context)
        )
        contextual_references = tuple(
            reference for use in contextual_uses for reference in use.classification.references
        )
        static_references = list(initial_semantic.references)
        static_issues = list(initial_semantic.issues)
        for use in contextual_uses:
            _remove_once(static_references, use.classification.references)
            _remove_once(static_issues, use.classification.issues)
        edges.extend(_block_edges(block, block_anchor, static_references))

        volatile = initial_semantic.volatile
        opaque = block.opaque or initial_semantic.opaque
        located_issues: list[tuple[FormulaIssue, FormulaAnchor]] = [
            (issue, block_anchor) for issue in static_issues
        ]
        for tile in _semantic_tiles(block, descriptor, context, contextual_references):
            tile_anchor = FormulaAnchor(
                descriptor.order,
                descriptor.name,
                tile.row_min,
                tile.col_min,
            )
            tile_block = replace(block, rect=tile)
            tile_formula = block.anchor_formula
            if tile_anchor.cell != block_anchor.cell:
                tile_formula = translate_a1_formula(
                    block.anchor_formula,
                    origin=block_anchor.cell,
                    target=tile_anchor.cell,
                    preserve_coordinate_spills=True,
                )
            tile_semantic = analyze_formula(
                tile_formula,
                anchor=tile_anchor,
                context=context,
            )
            for use in tile_semantic.reference_uses:
                if not _reference_use_requires_structured_context(use, context):
                    continue
                edges.extend(
                    _block_edges(
                        tile_block,
                        tile_anchor,
                        use.classification.references,
                    )
                )
                located_issues.extend((issue, tile_anchor) for issue in use.classification.issues)

            contextual_markers = tuple(
                reference for reference in tile_semantic.references if reference.via == "opaque:ref"
            )
            edges.extend(_block_edges(tile_block, tile_anchor, contextual_markers))

            dynamic_issues = tuple(
                issue for issue in tile_semantic.issues if issue.code == "I_DYNAMIC_REF"
            )
            dynamic_functions = {
                str(issue.related["function"])
                for issue in dynamic_issues
                if "function" in issue.related
            }
            dynamic_references = tuple(
                reference
                for reference in tile_semantic.references
                if reference.via.removeprefix("opaque:") in dynamic_functions
            )
            edges.extend(_block_edges(tile_block, tile_anchor, dynamic_references))
            located_issues.extend((issue, tile_anchor) for issue in dynamic_issues)
            volatile = volatile or tile_semantic.volatile
            opaque = opaque or tile_semantic.opaque
        blocks.append(
            replace(
                block,
                volatile=volatile,
                opaque=opaque,
            )
        )
        diagnostics.extend(
            _issue_diagnostics(
                descriptor.name,
                block,
                located_issues,
            )
        )

    final_blocks = tuple(blocks)
    for inconsistency in detect_inconsistent_formulas(patterns, regions, final_blocks):
        ref = make_cell_ref(inconsistency.row, inconsistency.col)
        diagnostics.append(
            FormulaDiagnostic(
                severity="warn",
                code="W_INCONSISTENT_FORMULA",
                row=inconsistency.row,
                col=inconsistency.col,
                ref=cell_symbol_id(descriptor.name, ref),
                message="Formula differs from the dominant pattern in its contiguous run.",
                related={
                    "dominantBlock": formula_block_symbol_id(
                        descriptor.name,
                        inconsistency.dominant_block_n,
                    ),
                    "expectedR1C1": inconsistency.expected_r1c1,
                },
            )
        )

    return SheetFormulaAnalysis(
        blocks=final_blocks,
        edges=tuple(sorted(set(edges), key=_edge_sort_key)),
        diagnostics=tuple(sorted(_deduplicate_diagnostics(diagnostics), key=_diagnostic_sort_key)),
    )


def _reference_use_requires_structured_context(
    use: ReferenceUse,
    context: ReferenceContext,
) -> bool:
    return any(
        _reference_requires_structured_context(reference, context)
        for reference in use.classification.references
    )


def _reference_requires_structured_context(
    reference: ExtractedReference,
    context: ReferenceContext,
) -> bool:
    if "[" in reference.token:
        requirement = structured_context_requirement(reference.token, context)
        return requirement.uses_current_table or bool(requirement.current_row_tables)
    return False


def _remove_once(items: list[_T], removals: Sequence[_T]) -> None:
    for removal in removals:
        try:
            items.remove(removal)
        except ValueError as exc:  # pragma: no cover - internal provenance invariant
            raise RuntimeError("formula reference-use provenance is inconsistent") from exc


def _semantic_tiles(
    block: FormulaBlock,
    descriptor: SheetDescriptor,
    context: ReferenceContext,
    references: Sequence[ExtractedReference],
) -> tuple[Rect, ...]:
    uses_current_table = False
    uses_current_table_row = False
    current_row_tables: set[TableBinding] = set()
    for reference in references:
        if "[" not in reference.token:
            continue
        requirement = structured_context_requirement(reference.token, context)
        uses_current_table = uses_current_table or requirement.uses_current_table
        uses_current_table_row = uses_current_table_row or requirement.uses_current_table_row
        current_row_tables.update(requirement.current_row_tables)
    if not uses_current_table and not current_row_tables:
        return (block.rect,)

    context_rectangles: list[tuple[str, Rect]] = []
    if uses_current_table:
        for binding in context.tables_intersecting(descriptor.order, block.rect):
            table_rect = context.table_rect(binding)
            intersection = _rect_intersection(block.rect, table_rect)
            if intersection is not None:
                context_rectangles.append((f"table:{binding.table.name.casefold()}", intersection))
            if uses_current_table_row:
                data_min = table_rect.row_min + binding.table.header_rows
                data_max = table_rect.row_max - binding.table.totals_rows
                if data_min <= data_max:
                    data_intersection = _rect_intersection(
                        block.rect,
                        Rect(
                            data_min,
                            data_max,
                            table_rect.col_min,
                            table_rect.col_max,
                        ),
                    )
                    if data_intersection is not None:
                        context_rectangles.append(
                            (
                                f"table-row:{binding.table.name.casefold()}",
                                data_intersection,
                            )
                        )
    for binding in sorted(
        current_row_tables,
        key=lambda item: (item.sheet_order, item.table.name.casefold()),
    ):
        if binding.sheet_order != descriptor.order:
            continue
        table_rect = context.table_rect(binding)
        data_min = table_rect.row_min + binding.table.header_rows
        data_max = table_rect.row_max - binding.table.totals_rows
        if data_min > data_max:
            continue
        intersection = _rect_intersection(
            block.rect,
            Rect(data_min, data_max, block.rect.col_min, block.rect.col_max),
        )
        if intersection is not None:
            context_rectangles.append((f"row:{binding.table.name.casefold()}", intersection))
    if not context_rectangles:
        return (block.rect,)
    return _partition_context_rectangles(block.rect, context_rectangles)


def _partition_context_rectangles(
    bounds: Rect,
    contexts: Sequence[tuple[str, Rect]],
) -> tuple[Rect, ...]:
    width = bounds.col_max - bounds.col_min + 1
    row_events: dict[int, list[tuple[int, str, int, int]]] = {}
    for label, raw_rect in contexts:
        rect = _rect_intersection(bounds, raw_rect)
        if rect is None:
            continue
        row_events.setdefault(rect.row_min, []).append((1, label, rect.col_min, rect.col_max))
        if rect.row_max < bounds.row_max:
            row_events.setdefault(rect.row_max + 1, []).append(
                (-1, label, rect.col_min, rect.col_max)
            )

    active_labels: list[dict[str, int]] = [{} for _ in range(width)]
    open_keys: list[tuple[str, ...]] = [()] * width
    open_rows = [bounds.row_min] * width
    completed: list[_SemanticTile] = []
    for row in sorted(row_events):
        touched: set[int] = set()
        for delta, label, col_min, col_max in sorted(row_events[row]):
            for col in range(col_min, col_max + 1):
                offset = col - bounds.col_min
                counts = active_labels[offset]
                count = counts.get(label, 0) + delta
                if count > 0:
                    counts[label] = count
                else:
                    counts.pop(label, None)
                touched.add(offset)

        changed: list[tuple[int, int, tuple[str, ...]]] = []
        for offset in sorted(touched):
            context_key = tuple(sorted(active_labels[offset]))
            if context_key == open_keys[offset]:
                continue
            if open_rows[offset] < row:
                changed.append((offset, open_rows[offset], open_keys[offset]))
            open_rows[offset] = row
            open_keys[offset] = context_key
        _close_column_tiles(
            completed,
            changed,
            row_max=row - 1,
            col_offset=bounds.col_min,
        )

    _close_column_tiles(
        completed,
        [
            (offset, open_rows[offset], open_keys[offset])
            for offset in range(width)
            if open_rows[offset] <= bounds.row_max
        ],
        row_max=bounds.row_max,
        col_offset=bounds.col_min,
    )
    return tuple(
        tile.rect
        for tile in sorted(
            completed,
            key=lambda item: (
                item.rect.row_min,
                item.rect.col_min,
                item.rect.row_max,
                item.rect.col_max,
                item.context_key,
            ),
        )
    )


def _close_column_tiles(
    completed: list[_SemanticTile],
    columns: Sequence[tuple[int, int, tuple[str, ...]]],
    *,
    row_max: int,
    col_offset: int,
) -> None:
    """Close adjacent column strips that share one vertical history."""
    if not columns:
        return
    first_offset, row_min, context_key = columns[0]
    last_offset = first_offset
    for offset, candidate_row_min, candidate_key in columns[1:]:
        if (
            offset == last_offset + 1
            and candidate_row_min == row_min
            and candidate_key == context_key
        ):
            last_offset = offset
            continue
        completed.append(
            _SemanticTile(
                Rect(
                    row_min,
                    row_max,
                    col_offset + first_offset,
                    col_offset + last_offset,
                ),
                context_key,
            )
        )
        first_offset = last_offset = offset
        row_min = candidate_row_min
        context_key = candidate_key
    completed.append(
        _SemanticTile(
            Rect(
                row_min,
                row_max,
                col_offset + first_offset,
                col_offset + last_offset,
            ),
            context_key,
        )
    )


def _rect_intersection(left: Rect, right: Rect) -> Rect | None:
    if not left.intersects(right):
        return None
    return Rect(
        max(left.row_min, right.row_min),
        min(left.row_max, right.row_max),
        max(left.col_min, right.col_min),
        min(left.col_max, right.col_max),
    )


def _block_edges(
    block: FormulaBlock,
    anchor: FormulaAnchor,
    references: Sequence[ExtractedReference],
) -> tuple[FormulaEdge, ...]:
    result: list[FormulaEdge] = []
    for reference in references:
        rect: Rect | None = None
        if reference.dst_sheet_order is not None and reference.geometry is not None:
            if reference.via == "spill":
                rect = resolve_reference(reference.geometry, anchor.cell)
            else:
                geometries = reference.extrusion_geometries or (reference.geometry,)
                component_rects = tuple(
                    extrude_reference(geometry, block.rect) for geometry in geometries
                )
                rect = Rect(
                    min(component.row_min for component in component_rects),
                    max(component.row_max for component in component_rects),
                    min(component.col_min for component in component_rects),
                    max(component.col_max for component in component_rects),
                )
        elif reference.dst_sheet_order is not None or reference.geometry is not None:
            raise ValueError("extracted reference has an incomplete destination")
        result.append(
            FormulaEdge(
                source_block_n=block.n,
                dst_sheet_order=reference.dst_sheet_order,
                rect=rect,
                via=reference.via,
            )
        )
    return tuple(result)


def _issue_diagnostics(
    sheet_name: str,
    block: FormulaBlock,
    located_issues: Sequence[tuple[FormulaIssue, FormulaAnchor]],
) -> tuple[FormulaDiagnostic, ...]:
    block_id = formula_block_symbol_id(sheet_name, block.n)
    result: list[FormulaDiagnostic] = []
    unique: dict[tuple[object, ...], tuple[FormulaIssue, FormulaAnchor]] = {}
    for issue, anchor in located_issues:
        if issue.code not in _P3_ISSUE_CODES:
            continue
        key = (
            issue.severity,
            issue.code,
            issue.message,
            tuple(sorted((name, repr(value)) for name, value in issue.related.items())),
        )
        existing = unique.get(key)
        if existing is None or (anchor.row, anchor.col) < (
            existing[1].row,
            existing[1].col,
        ):
            unique[key] = (issue, anchor)
    for issue, anchor in sorted(
        unique.values(),
        key=lambda item: (
            item[1].row,
            item[1].col,
            item[0].code,
            item[0].message,
        ),
    ):
        cell_ref = make_cell_ref(anchor.row, anchor.col)
        related = dict(issue.related)
        related["block"] = block_id
        result.append(
            FormulaDiagnostic(
                severity=issue.severity,
                code=issue.code,
                row=anchor.row,
                col=anchor.col,
                ref=cell_symbol_id(sheet_name, cell_ref),
                message=issue.message,
                related=related,
            )
        )
    return tuple(result)


def _deduplicate_diagnostics(
    diagnostics: Sequence[FormulaDiagnostic],
) -> tuple[FormulaDiagnostic, ...]:
    result: list[FormulaDiagnostic] = []
    seen: set[tuple[object, ...]] = set()
    for diagnostic in diagnostics:
        key = (
            diagnostic.severity,
            diagnostic.code,
            diagnostic.row,
            diagnostic.col,
            diagnostic.ref,
            diagnostic.message,
            tuple(sorted((key, repr(value)) for key, value in diagnostic.related.items())),
        )
        if key not in seen:
            seen.add(key)
            result.append(diagnostic)
    return tuple(result)


def _edge_sort_key(edge: FormulaEdge) -> tuple[object, ...]:
    rect = edge.rect
    return (
        edge.source_block_n,
        -1 if edge.dst_sheet_order is None else edge.dst_sheet_order,
        -1 if rect is None else rect.row_min,
        -1 if rect is None else rect.col_min,
        -1 if rect is None else rect.row_max,
        -1 if rect is None else rect.col_max,
        edge.via,
    )


def _diagnostic_sort_key(diagnostic: FormulaDiagnostic) -> tuple[object, ...]:
    return (
        diagnostic.row,
        diagnostic.col,
        diagnostic.code,
        diagnostic.message,
        tuple(sorted((key, repr(value)) for key, value in diagnostic.related.items())),
    )


__all__ = [
    "FormulaDiagnostic",
    "FormulaEdge",
    "SheetFormulaAnalysis",
    "analyze_sheet_formulas",
]
