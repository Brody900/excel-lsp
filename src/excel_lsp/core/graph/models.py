"""Immutable dependency-graph query contracts."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from excel_lsp.core.models import Rect
from excel_lsp.core.parse.coordinates import make_cell_ref

GraphTargetKind = Literal["fblock", "cell", "range", "opaque"]
GraphDirection = Literal["precedents", "dependents"]


@dataclass(frozen=True, slots=True)
class GraphArea:
    """One worksheet rectangle without any cell-level expansion."""

    sheet_id: int
    sheet: str
    rect: Rect

    def __post_init__(self) -> None:
        if type(self.sheet_id) is not int or self.sheet_id < 1:
            raise ValueError("sheet_id must be positive")
        if not self.sheet:
            raise ValueError("sheet must not be empty")

    @property
    def ref(self) -> str:
        """Return a deterministic sheet-qualified A1 reference."""
        start = make_cell_ref(self.rect.row_min, self.rect.col_min)
        end = make_cell_ref(self.rect.row_max, self.rect.col_max)
        local_ref = start if start == end else f"{start}:{end}"
        if self.sheet.replace("_", "a").isalnum() and not self.sheet[0].isdigit():
            return f"{self.sheet}!{local_ref}"
        return f"'{self.sheet.replace(chr(39), chr(39) * 2)}'!{local_ref}"


@dataclass(frozen=True, slots=True)
class GraphTarget:
    """A semantic graph node: block, singleton cell, range, or opaque sink."""

    kind: GraphTargetKind
    symbol: str | None
    ref: str | None
    area: GraphArea | None

    def __post_init__(self) -> None:
        if self.kind == "opaque":
            if self.area is not None or self.symbol is not None or not self.ref:
                raise ValueError("opaque targets require only a descriptive ref")
            return
        if self.area is None:
            raise ValueError("concrete graph targets require an area")
        if self.kind in {"fblock", "cell"} and not self.symbol:
            raise ValueError("block and cell targets require a symbol")
        if self.kind == "range" and self.symbol is not None:
            raise ValueError("range targets do not have frozen symbol ids")
        if self.ref != self.area.ref:
            raise ValueError("concrete target ref must match its area")

    @property
    def label(self) -> str:
        """Return the preferred public node label."""
        return self.symbol or self.ref or ""


@dataclass(frozen=True, slots=True)
class GraphHop:
    """One directed graph edge and its semantic reason."""

    target: GraphTarget
    via: str


@dataclass(frozen=True, slots=True)
class TraceNode:
    """One immutable node in a bounded dependency trace tree.

    ``child_count`` is exact when the containing result is not truncated.  At
    the truncation boundary it is a bounded lower-bound witness, so a value
    greater than ``len(children)`` proves that at least one child was omitted.
    """

    target: GraphTarget
    via: str | None
    children: tuple[TraceNode, ...] = ()
    child_count: int = 0


@dataclass(frozen=True, slots=True)
class TraceResult:
    """A bounded dependency trace and exact emitted counts."""

    direction: GraphDirection
    root: TraceNode
    node_count: int
    edge_count: int
    truncated: bool


@dataclass(frozen=True, slots=True)
class PathStep:
    """One source/reference or block step in a dependent-direction path."""

    symbol: str
    via: str | None


@dataclass(frozen=True, slots=True)
class PathResult:
    """Bounded shortest dependent paths between two worksheet areas."""

    connected: bool
    paths: tuple[tuple[PathStep, ...], ...]
    truncated: bool = False


__all__ = [
    "GraphArea",
    "GraphDirection",
    "GraphHop",
    "GraphTarget",
    "GraphTargetKind",
    "PathResult",
    "PathStep",
    "TraceNode",
    "TraceResult",
]
