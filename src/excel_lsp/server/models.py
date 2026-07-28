"""Pydantic contracts used by the MCP write surface."""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict


class WriteCellInput(BaseModel):
    """One qualified cell write with exactly one operation."""

    model_config = ConfigDict(extra="forbid", strict=True)

    ref: str
    value: int | float | str | bool | None = None
    formula: str | None = None


class ToolEnvelope(BaseModel):
    """Permissive structured result model shared by heterogeneous tools."""

    model_config = ConfigDict(extra="allow")

    error: dict[str, Any] | None = None


__all__ = ["ToolEnvelope", "WriteCellInput"]
