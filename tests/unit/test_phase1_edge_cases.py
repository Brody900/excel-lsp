"""Phase-one spatial-index and workbook-lifecycle edge contracts."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any, cast

import pytest
from openpyxl import Workbook

import excel_lsp.core.index.lifecycle as lifecycle
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.index import IndexStore, index_workbook
from excel_lsp.core.index.edges import EdgeStore
from excel_lsp.core.models import Rect


class _RtreeFailureProxy:
    """Delegate SQLite operations except for deterministic R*Tree creation failure."""

    def __init__(self, connection: sqlite3.Connection, message: str) -> None:
        self.connection = connection
        self.message = message

    def execute(self, sql: str, parameters: Any = ()) -> sqlite3.Cursor:
        if "CREATE VIRTUAL TABLE edge_rtree" in sql:
            raise sqlite3.OperationalError(self.message)
        return self.connection.execute(sql, parameters)

    def executescript(self, sql: str) -> sqlite3.Cursor:
        return self.connection.executescript(sql)


def _new_workbook(path: Path) -> Path:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = 1
    workbook.save(path)
    workbook.close()
    return path


def test_edge_store_requires_initialized_schema() -> None:
    connection = sqlite3.connect(":memory:")
    try:
        with pytest.raises(RuntimeError, match="schema has not been initialized"):
            EdgeStore(connection)
    finally:
        connection.close()


def test_edge_schema_falls_back_only_when_rtree_module_is_unavailable() -> None:
    connection = sqlite3.connect(":memory:", isolation_level=None)
    proxy = cast(sqlite3.Connection, _RtreeFailureProxy(connection, "no such module: rtree"))
    try:
        assert EdgeStore.ensure_schema(proxy) == "interval"
        assert EdgeStore.ensure_schema(connection) == "interval"

        edges = EdgeStore(connection)
        edges.insert(1, 4, Rect(1, 1_048_576, 2, 2))
        assert edges.query_point(4, 1_048_576, 2) == (1,)
        edges.clear()
        assert edges.query_point(4, 1, 2) == ()
    finally:
        connection.close()


@pytest.mark.parametrize("message", ["disk I/O error", "no such module: geometry"])
def test_edge_schema_does_not_hide_unrelated_sqlite_failures(message: str) -> None:
    connection = sqlite3.connect(":memory:", isolation_level=None)
    proxy = cast(sqlite3.Connection, _RtreeFailureProxy(connection, message))
    try:
        with pytest.raises(sqlite3.OperationalError, match=message):
            EdgeStore.ensure_schema(proxy)
        assert EdgeStore.INTERVAL_TABLE not in {
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
        }
    finally:
        connection.close()


@pytest.mark.parametrize(
    ("kind", "expected_code"),
    [
        ("missing", ErrorCode.NOT_FOUND),
        ("directory", ErrorCode.NOT_FOUND),
        ("xls", ErrorCode.UNSUPPORTED_FORMAT),
    ],
)
def test_index_workbook_uses_structured_path_errors(
    tmp_path: Path, kind: str, expected_code: ErrorCode
) -> None:
    candidate = tmp_path / ("missing.xlsx" if kind == "missing" else f"candidate.{kind}")
    if kind == "directory":
        candidate.mkdir()
    elif kind == "xls":
        candidate.write_bytes(b"legacy workbook marker")

    with pytest.raises(ExcelLSPError) as caught:
        index_workbook(candidate, index_dir=tmp_path / "indexes")

    assert caught.value.code is expected_code
    assert caught.value.as_dict()["error"]["code"] == expected_code.value


def test_index_workbook_reports_permission_failures_as_locked(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    candidate = tmp_path / "locked.xlsx"
    candidate.write_bytes(b"content is never parsed")
    resolved_candidate = candidate.resolve()
    original_stat = Path.stat

    def denied_stat(path: Path, *, follow_symlinks: bool = True) -> Any:
        if path == resolved_candidate:
            raise PermissionError("sharing violation")
        return original_stat(path, follow_symlinks=follow_symlinks)

    monkeypatch.setattr(Path, "stat", denied_stat)

    with pytest.raises(ExcelLSPError) as caught:
        index_workbook(candidate, index_dir=tmp_path / "indexes")

    assert caught.value.code is ErrorCode.LOCKED
    assert caught.value.hint == "Close the application holding the file and retry."


def test_invalid_fast_path_metadata_is_repaired_by_hash_noop(tmp_path: Path) -> None:
    path = _new_workbook(tmp_path / "metadata.xlsx")
    first = index_workbook(path, index_dir=tmp_path / "indexes")
    with IndexStore(first.index_path) as store:
        store.set_meta("mtime_ns", "not-an-integer")

    second = index_workbook(path, index_dir=tmp_path / "indexes")

    assert second.changed is False
    assert second.generation == first.generation
    assert second.reindexed_sheets == ()
    with IndexStore(second.index_path) as store:
        assert store.get_meta("mtime_ns") == str(path.stat().st_mtime_ns)
        assert store.generation == first.generation


def test_repeated_torn_save_detection_retries_once_then_reports_corrupt(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = _new_workbook(tmp_path / "changing.xlsx")
    comparisons = 0

    def always_changed(_left: Any, _right: Any) -> bool:
        nonlocal comparisons
        comparisons += 1
        return False

    monkeypatch.setattr(lifecycle, "_same_stat", always_changed)

    with pytest.raises(ExcelLSPError) as caught:
        index_workbook(path, index_dir=tmp_path / "indexes")

    assert comparisons == 2
    assert caught.value.code is ErrorCode.CORRUPT
    assert caught.value.hint == "Wait for Excel to finish saving, then retry refresh."
