"""Single public value-normalization boundary."""

from __future__ import annotations

from datetime import date, datetime, time
from typing import TypeAlias

JsonScalar: TypeAlias = str | int | float | bool | None


def normalize_value(value: object) -> JsonScalar:
    """Convert a parsed value to the JSON scalar used by every public surface."""
    if value is None or isinstance(value, (str, bool, int, float)):
        return value
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    raise TypeError(f"unsupported cell value type: {type(value).__name__}")
