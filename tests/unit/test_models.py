"""Tests for shared Phase 1 contracts."""

from datetime import datetime

import pytest

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.models import Rect
from excel_lsp.core.values import normalize_value


def test_rect_intersection_is_inclusive() -> None:
    assert Rect(1, 2, 1, 2).intersects(Rect(2, 3, 2, 3))
    assert not Rect(1, 1, 1, 1).intersects(Rect(2, 2, 1, 1))


def test_rect_rejects_out_of_bounds_coordinates() -> None:
    with pytest.raises(ValueError, match="one-based"):
        Rect(0, 1, 1, 1)


def test_normalize_value_uses_iso_dates() -> None:
    assert normalize_value(datetime(2026, 7, 15, 12, 30)) == "2026-07-15T12:30:00"


def test_structured_error_omits_absent_optional_fields() -> None:
    error = ExcelLSPError(ErrorCode.NOT_FOUND, "missing")

    assert error.as_dict() == {"error": {"code": "E_NOT_FOUND", "message": "missing"}}
