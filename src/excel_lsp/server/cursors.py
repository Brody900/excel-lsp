"""Opaque, generation-bound pagination cursors."""

from __future__ import annotations

import base64
import hashlib
import json
from collections.abc import Mapping
from typing import Any

from excel_lsp.core.errors import ErrorCode, ExcelLSPError


def parameter_hash(parameters: Mapping[str, object]) -> str:
    """Return a stable digest for the query parameters bound to a cursor."""
    encoded = json.dumps(
        parameters,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def encode_cursor(*, tool: str, params_hash: str, offset: int, generation: int) -> str:
    """Encode the frozen cursor payload as URL-safe base64."""
    payload = {"tool": tool, "params_hash": params_hash, "offset": offset, "gen": generation}
    raw = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def decode_cursor(
    cursor: str,
    *,
    tool: str,
    params_hash: str,
    generation: int,
) -> int:
    """Validate a cursor and return its nonnegative page offset."""
    error = ExcelLSPError(
        ErrorCode.STALE_CURSOR,
        "The pagination cursor is invalid or no longer matches this index generation.",
        hint="Re-issue the original query without a cursor.",
    )
    try:
        if not cursor or len(cursor) > 2_048:
            raise ValueError("invalid cursor length")
        padded = cursor + "=" * (-len(cursor) % 4)
        decoded: Any = json.loads(base64.urlsafe_b64decode(padded.encode("ascii")))
        if not isinstance(decoded, dict) or set(decoded) != {
            "tool",
            "params_hash",
            "offset",
            "gen",
        }:
            raise ValueError("invalid cursor shape")
        if (
            decoded["tool"] != tool
            or decoded["params_hash"] != params_hash
            or decoded["gen"] != generation
            or type(decoded["offset"]) is not int
            or decoded["offset"] < 0
        ):
            raise ValueError("cursor binding mismatch")
    except (UnicodeError, ValueError, TypeError, json.JSONDecodeError) as exc:
        raise error from exc
    return int(decoded["offset"])


__all__ = ["decode_cursor", "encode_cursor", "parameter_hash"]
