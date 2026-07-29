"""Canonical errors shared by the core library, CLI, and MCP server."""

from __future__ import annotations

from enum import StrEnum
from typing import Any


class ErrorCode(StrEnum):
    """Stable agent-facing error codes from the public tool contract."""

    NOT_FOUND = "E_NOT_FOUND"
    UNSUPPORTED_FORMAT = "E_UNSUPPORTED_FORMAT"
    ENCRYPTED = "E_ENCRYPTED"
    LOCKED = "E_LOCKED"
    OPEN_IN_EXCEL = "E_OPEN_IN_EXCEL"
    CONFLICT = "E_CONFLICT"
    CORRUPT = "E_CORRUPT"
    STALE_CURSOR = "E_STALE_CURSOR"
    INVALID_REF = "E_INVALID_REF"
    UNKNOWN_SYMBOL = "E_UNKNOWN_SYMBOL"
    ARRAY_FORMULA = "E_ARRAY_FORMULA"
    INVALID_VALUE = "E_INVALID_VALUE"
    PATH_DENIED = "E_PATH_DENIED"
    INTERNAL = "E_INTERNAL"


class ExcelLSPError(Exception):
    """Expected failure that is safe to return to an agent."""

    def __init__(
        self,
        code: ErrorCode,
        message: str,
        *,
        hint: str | None = None,
        details: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint
        self.details = details

    def as_dict(self) -> dict[str, Any]:
        """Return the canonical structured-error envelope."""
        payload: dict[str, Any] = {"code": self.code.value, "message": self.message}
        if self.hint:
            payload["hint"] = self.hint
        if self.details:
            payload["details"] = self.details
        return {"error": payload}
