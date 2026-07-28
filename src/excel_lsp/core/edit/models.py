"""Typed contracts for surgical workbook edits."""

from __future__ import annotations

from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from typing import TypeAlias

WriteScalar: TypeAlias = int | float | str | bool | None


class CellEditKind(StrEnum):
    """The mutually exclusive worksheet-cell edit modes."""

    VALUE = "value"
    FORMULA = "formula"


@dataclass(frozen=True, slots=True)
class CellEdit:
    """One qualified A1 cell edit."""

    sheet: str
    ref: str
    kind: CellEditKind
    payload: WriteScalar

    @classmethod
    def value(cls, sheet: str, ref: str, value: WriteScalar) -> CellEdit:
        """Construct a value edit, including an explicit null write."""
        return cls(sheet=sheet, ref=ref, kind=CellEditKind.VALUE, payload=value)

    @classmethod
    def formula(cls, sheet: str, ref: str, formula: str) -> CellEdit:
        """Construct a formula edit."""
        return cls(sheet=sheet, ref=ref, kind=CellEditKind.FORMULA, payload=formula)


@dataclass(frozen=True, slots=True)
class PatchedCell:
    """One cell whose worksheet XML changed, including shared-group expansion."""

    sheet: str
    ref: str
    requested: bool


@dataclass(frozen=True, slots=True)
class PatchResult:
    """Observable outcome of one atomic surgical workbook replacement."""

    path: Path
    workbook_hash_before: str
    workbook_hash_after: str
    modified_parts: tuple[str, ...]
    deleted_parts: tuple[str, ...]
    patched_cells: tuple[PatchedCell, ...]


@dataclass(frozen=True, slots=True)
class EditResult:
    """Workbook-and-index outcome returned by the core write service."""

    patch: PatchResult
    generation: int
    stale_blocks: int
    direct_index_patch: bool


@dataclass(frozen=True, slots=True)
class ColumnFormulaResult:
    """Outcome of filling one indexed semantic column with formulas."""

    edit: EditResult
    formula_block: str
    cells_written: int


__all__ = [
    "CellEdit",
    "CellEditKind",
    "ColumnFormulaResult",
    "EditResult",
    "PatchResult",
    "PatchedCell",
    "WriteScalar",
]
