"""Typed contracts shared by OOXML parsing and index storage."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from datetime import date, datetime, time
from types import MappingProxyType
from typing import Literal, TypeAlias

CellScalar: TypeAlias = int | float | str | bool | date | datetime | time | None
CellValueType: TypeAlias = Literal["number", "string", "bool", "error", "date", "blank"]
FormulaKind: TypeAlias = Literal["normal", "shared", "array", "dataTable"]
SheetKind: TypeAlias = Literal["worksheet", "chartsheet", "macro", "dialog"]
Visibility: TypeAlias = Literal["visible", "hidden", "veryHidden"]
DefinedNameKind: TypeAlias = Literal["range", "multi_range", "constant", "formula", "lambda"]
CalculationMode: TypeAlias = Literal["manual", "auto", "autoNoTable"]
ReferenceMode: TypeAlias = Literal["A1", "R1C1"]


@dataclass(frozen=True, slots=True)
class Rect:
    """Inclusive, one-based worksheet rectangle."""

    row_min: int
    row_max: int
    col_min: int
    col_max: int

    def __post_init__(self) -> None:
        if self.row_min < 1 or self.col_min < 1:
            raise ValueError("worksheet coordinates are one-based")
        if self.row_max < self.row_min or self.col_max < self.col_min:
            raise ValueError("rectangle maxima must not precede minima")
        if self.row_max > 1_048_576 or self.col_max > 16_384:
            raise ValueError("rectangle exceeds Excel worksheet bounds")

    def intersects(self, other: Rect) -> bool:
        """Return whether two inclusive rectangles overlap."""
        return not (
            self.row_max < other.row_min
            or other.row_max < self.row_min
            or self.col_max < other.col_min
            or other.col_max < self.col_min
        )


@dataclass(frozen=True, slots=True)
class DataTableFormulaInfo:
    """Metadata for an OOXML What-If Data Table formula span."""

    ref: str
    rect: Rect
    input_cell_1: str | None = None
    input_cell_2: str | None = None
    is_2d: bool = False
    row_oriented: bool = False
    calculate_always: bool = False
    deleted_row_input: bool = False
    deleted_column_input: bool = False


@dataclass(frozen=True, slots=True)
class CellRecord:
    """One non-empty OOXML cell emitted by the streaming parser."""

    ref: str
    row: int
    col: int
    value: CellScalar
    value_type: CellValueType
    formula: str | None = None
    style_idx: int = 0
    formula_kind: FormulaKind | None = None
    shared_index: int | None = None
    array_ref: str | None = None
    data_table: DataTableFormulaInfo | None = None


@dataclass(frozen=True, slots=True)
class SheetDescriptor:
    """Workbook-order sheet metadata resolved through workbook relationships."""

    order: int
    name: str
    sheet_id: int
    rel_id: str
    xml_part: str
    kind: SheetKind
    visibility: Visibility = "visible"
    related_parts: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class TableInfo:
    """Native ListObject metadata parsed from a table part."""

    name: str
    display_name: str
    ref: str
    header_rows: int
    totals_rows: int
    columns: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class DataValidationInfo:
    """One validation constraint applied to one rectangle."""

    rect: Rect
    validation_type: str | None
    operator: str | None
    formula1: str | None
    formula2: str | None
    allow_blank: bool


@dataclass(frozen=True, slots=True)
class NameArea:
    """One concrete worksheet area belonging to a defined name."""

    sheet_name: str
    rect: Rect


@dataclass(frozen=True, slots=True)
class DefinedName:
    """Workbook- or sheet-scoped defined name."""

    name: str
    refers_to: str
    scope_sheet_order: int | None
    kind: DefinedNameKind
    is_builtin: bool
    areas: tuple[NameArea, ...] = ()


@dataclass(frozen=True, slots=True)
class SheetParseSummary:
    """Small metadata result returned after a streaming sheet parse."""

    descriptor: SheetDescriptor
    part_hash: str
    max_row: int
    max_col: int
    cell_count: int
    dimension_ref: str | None = None
    merges: tuple[Rect, ...] = ()
    validations: tuple[DataValidationInfo, ...] = ()
    tables: tuple[TableInfo, ...] = ()
    array_formulas: tuple[Rect, ...] = ()
    data_tables: tuple[DataTableFormulaInfo, ...] = ()


@dataclass(frozen=True, slots=True)
class CalculationProperties:
    """Typed attributes from workbook.xml ``calcPr`` when present."""

    calc_id: int | None = None
    calc_mode: CalculationMode | None = None
    full_calc_on_load: bool | None = None
    ref_mode: ReferenceMode | None = None
    iterate: bool | None = None
    iterate_count: int | None = None
    iterate_delta: float | None = None
    full_precision: bool | None = None
    calc_completed: bool | None = None
    calc_on_save: bool | None = None
    concurrent_calc: bool | None = None
    concurrent_manual_count: int | None = None
    force_full_calc: bool | None = None


@dataclass(frozen=True, slots=True)
class WorkbookMetadata:
    """Workbook-level OOXML metadata required by downstream modules."""

    path: str
    date1904: bool
    sheets: tuple[SheetDescriptor, ...]
    defined_names: tuple[DefinedName, ...]
    calculation: CalculationProperties = field(default_factory=CalculationProperties)
    external_links: Mapping[int, str] = field(default_factory=lambda: MappingProxyType({}))
    has_vba: bool = False


@dataclass(frozen=True, slots=True)
class PackageHashes:
    """Whole-file and selected OOXML-part SHA-256 values."""

    whole_file: str
    parts: Mapping[str, str]


@dataclass(frozen=True, slots=True)
class IndexUpdate:
    """Observable result of opening or refreshing a workbook index."""

    workbook_path: str
    index_path: str
    generation: int
    changed: bool
    reindexed_sheets: tuple[str, ...]
