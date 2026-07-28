"""Focused SQLite store, canonical export, and spatial-backend tests."""

from __future__ import annotations

import sqlite3
from collections.abc import Iterator
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from random import Random
from threading import Barrier
from time import perf_counter
from typing import Any, cast

import pytest

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.graph import GraphArea
from excel_lsp.core.index.edges import (
    EdgeDirection,
    EdgeSchemaError,
    EdgeStore,
    RankedEdge,
    canonical_ranked_edges,
)
from excel_lsp.core.index.schema import SCHEMA_VERSION
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import (
    CellRecord,
    DataValidationInfo,
    Rect,
    SheetDescriptor,
    SheetParseSummary,
    TableInfo,
)
from excel_lsp.core.regions import RegionOptions


def _descriptor(name: str = "Data", *, order: int = 0, kind: str = "worksheet") -> SheetDescriptor:
    return SheetDescriptor(
        order=order,
        name=name,
        sheet_id=order + 10,
        rel_id=f"rId{order + 1}",
        xml_part=(
            f"xl/worksheets/sheet{order + 1}.xml"
            if kind == "worksheet"
            else f"xl/chartsheets/sheet{order + 1}.xml"
        ),
        kind=kind,  # type: ignore[arg-type]
    )


def test_store_configures_sqlite_and_creates_frozen_schema(tmp_path: Path) -> None:
    database = tmp_path / "book.xlsp.db"

    with IndexStore(database) as store:
        tables = {
            str(row[0])
            for row in store.connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
        }
        assert {
            "meta",
            "sheets",
            "regions",
            "list_objects",
            "list_object_columns",
            "columns",
            "fblocks",
            "defined_names",
            "name_areas",
            "validations",
            "edges",
            "diagnostics",
            "staleness",
            "cells",
            "package_parts",
            "graph_spatial_state",
            "graph_rank_keys",
            store.edge_store.table_name,
            store.edge_store.source_table_name,
        } <= tables
        assert store.get_meta("schema_version") == SCHEMA_VERSION
        assert store.connection.execute("PRAGMA journal_mode").fetchone()[0] == "wal"
        assert store.connection.execute("PRAGMA busy_timeout").fetchone()[0] == 5_000
        assert store.connection.execute("PRAGMA synchronous").fetchone()[0] == 1


def test_streaming_sheet_replace_normalizes_values_and_stores_formulas(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor()
    cells = (
        CellRecord("A1", 1, 1, datetime(2026, 7, 15, 9, 30), "date"),
        CellRecord("B1", 1, 2, True, "bool"),
        CellRecord("C1", 1, 3, 3, "number", formula="=1+2"),
    )
    validation = DataValidationInfo(
        rect=Rect(2, 10, 1, 1),
        validation_type="whole",
        operator="between",
        formula1="1",
        formula2="10",
        allow_blank=True,
    )

    with IndexStore(tmp_path / "stream.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))

        def parse(on_cell: object) -> SheetParseSummary:
            consumer = on_cell  # retain a clear assertion site for callback streaming
            assert callable(consumer)
            for cell in cells:
                consumer(cell)
            return SheetParseSummary(
                descriptor=descriptor,
                part_hash="sheet-hash",
                max_row=10,
                max_col=3,
                cell_count=3,
                validations=(validation,),
            )

        summary = store.replace_sheet(descriptor, parse)  # type: ignore[arg-type]
        rows = store.connection.execute(
            "SELECT ref, value, value_type, formula FROM cells ORDER BY row, col"
        ).fetchall()
        assert summary.cell_count == 3
        assert [tuple(row) for row in rows] == [
            ("A1", "2026-07-15T09:30:00", "date", None),
            ("B1", 1, "bool", None),
            ("C1", 3, "number", "=1+2"),
        ]
        assert store.connection.execute("SELECT COUNT(*) FROM validations").fetchone()[0] == 1
        assert store.generation == 1


def test_sheet_replace_persists_complete_list_object_catalog(tmp_path: Path) -> None:
    descriptor = _descriptor()
    table = TableInfo(
        name="SalesTable",
        display_name="SalesTable",
        ref="B2:D8",
        header_rows=1,
        totals_rows=1,
        columns=("Item", "Net Sales", "Tax"),
    )

    with IndexStore(tmp_path / "tables.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))

        def parse(on_cell: object) -> SheetParseSummary:
            assert callable(on_cell)
            for ref, col, value in zip(
                ("B2", "C2", "D2"),
                range(2, 5),
                table.columns,
                strict=True,
            ):
                on_cell(CellRecord(ref, 2, col, value, "string"))
            return SheetParseSummary(
                descriptor,
                "table-part-hash",
                8,
                4,
                len(table.columns),
                tables=(table,),
            )

        store.replace_sheet(descriptor, parse)  # type: ignore[arg-type]

        catalog = store.connection.execute(
            """
            SELECT s.name, t.name, t.lookup_name, t.display_name,
                   t.row_min, t.row_max, t.col_min, t.col_max,
                   t.header_rows, t.totals_rows
            FROM list_objects AS t
            JOIN sheets AS s ON s.id = t.sheet_id
            """
        ).fetchone()
        assert tuple(catalog) == (
            "Data",
            "SalesTable",
            "salestable",
            "SalesTable",
            2,
            8,
            2,
            4,
            1,
            1,
        )
        columns = store.connection.execute(
            """
            SELECT idx, name, lookup_name
            FROM list_object_columns
            ORDER BY idx
            """
        ).fetchall()
        assert tuple(map(tuple, columns)) == (
            (0, "Item", "item"),
            (1, "Net Sales", "net sales"),
            (2, "Tax", "tax"),
        )
        assert store.canonical_export()["list_objects"] == (
            ("Data", "SalesTable", "SalesTable", 2, 8, 2, 4, 1, 1),
        )


def test_nonworksheet_sheet_replacement_keeps_catalog_row_without_cells(tmp_path: Path) -> None:
    descriptor = _descriptor("Chart", kind="chartsheet")

    with IndexStore(tmp_path / "chart.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))
        store.replace_sheet(
            descriptor,
            lambda _on_cell: SheetParseSummary(
                descriptor=descriptor,
                part_hash="chart-hash",
                max_row=0,
                max_col=0,
                cell_count=0,
            ),
        )
        row = store.connection.execute("SELECT name, kind, max_row, max_col FROM sheets").fetchone()
        assert tuple(row) == ("Chart", "chartsheet", 0, 0)
        assert store.connection.execute("SELECT COUNT(*) FROM cells").fetchone()[0] == 0


def test_schema_version_mismatch_rebuilds_and_preserves_monotonic_generation(
    tmp_path: Path,
) -> None:
    database = tmp_path / "old.xlsp.db"
    with IndexStore(database) as store:
        store.set_meta("generation", 9)
        store.set_meta("schema_version", "old")
        store.connection.execute(
            "INSERT INTO package_parts VALUES ('xl/workbook.xml', 'old', 'workbook')"
        )

    with IndexStore(database) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == 10
        assert rebuilt.get_meta("schema_version") == SCHEMA_VERSION
        assert rebuilt.connection.execute("SELECT COUNT(*) FROM package_parts").fetchone()[0] == 0


def test_wal_allows_reader_during_uncommitted_writer_transaction(tmp_path: Path) -> None:
    database = tmp_path / "concurrent.xlsp.db"
    with IndexStore(database) as writer, IndexStore(database) as reader:
        writer.set_meta("marker", "old")
        with writer.transaction():
            writer.set_meta("marker", "new")
            assert reader.get_meta("marker") == "old"
        assert reader.get_meta("marker") == "new"


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize(
    ("failure_point", "rollback_fails"),
    [
        pytest.param("before-begin", False, id="before-begin"),
        pytest.param("after-begin", False, id="after-begin"),
        pytest.param("after-begin", True, id="after-begin-rollback-fails"),
    ],
)
@pytest.mark.parametrize("close_timing", ["before", "after"])
def test_constructor_failure_conclusively_closes_native_connection(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    journal_mode: str,
    failure_point: str,
    rollback_fails: bool,
    close_timing: str,
) -> None:
    import excel_lsp.core.index.store as store_module

    database = tmp_path / (
        f"constructor-{prefer_rtree}-{journal_mode}-{failure_point}-"
        f"{rollback_fails}-{close_timing}.xlsp.db"
    )
    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_CloseFailureConnection,
    )
    assert isinstance(connection, _CloseFailureConnection)
    journal_row = connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()
    assert journal_row is not None
    assert str(journal_row[0]) == journal_mode
    connection.row_factory = sqlite3.Row
    connection.close_timing = close_timing
    connection.close_marker = RuntimeError(f"constructor close {close_timing} effect")
    connection.state_spoof = "false-positive" if close_timing == "before" else "false-negative"
    connection.rollback_marker = RuntimeError("constructor virtual rollback failed")
    connection.rollback_fails = rollback_fails

    prior_error = LookupError("prior constructor causal evidence")
    primary_error = ValueError(f"constructor initialization failed {failure_point}")
    primary_error.__cause__ = prior_error
    primary_error.__suppress_context__ = True
    constructed: list[IndexStore] = []

    def open_instrumented_connection(_path: Path) -> sqlite3.Connection:
        return connection

    def fail_initialization(self: IndexStore, *, prefer_rtree: bool) -> None:
        del prefer_rtree
        constructed.append(self)
        if failure_point == "after-begin":
            sqlite3.Connection.execute(connection, "BEGIN IMMEDIATE")
        raise primary_error

    with monkeypatch.context() as patch:
        patch.setattr(store_module, "_open_index_connection", open_instrumented_connection)
        patch.setattr(IndexStore, "_initialize_schema", fail_initialization)
        with pytest.raises(ValueError) as captured:
            IndexStore(database, prefer_rtree=prefer_rtree)

    assert captured.value is primary_error
    assert len(constructed) == 1
    assert constructed[0]._closed is True
    assert not hasattr(constructed[0], "edge_store")
    assert connection.close_calls == 1
    assert connection.rollback_calls == (1 if failure_point == "after-begin" else 0)
    _assert_exception_identity_once(primary_error, prior_error)
    _assert_exception_identity_once(primary_error, connection.close_marker)
    if rollback_fails:
        _assert_exception_identity_once(primary_error, connection.rollback_marker)
    else:
        _assert_exception_identity_absent(primary_error, connection.rollback_marker)
    _assert_acyclic_exception_graph(primary_error)
    notes = getattr(primary_error, "__notes__", ())
    assert (
        notes.count(
            "IndexStore initialization cleanup also failed; the original error remains primary."
        )
        == 1
    )
    with pytest.raises(sqlite3.ProgrammingError, match="closed"):
        sqlite3.Connection.execute(connection, "SELECT 1")
    _assert_writer_unblocked(database)


@pytest.mark.parametrize("close_timing", ["before", "after"])
def test_constructor_owns_connection_before_capability_construction(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    close_timing: str,
) -> None:
    import excel_lsp.core.index.store as store_module

    database = tmp_path / f"constructor-capability-{close_timing}.xlsp.db"
    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_CloseFailureConnection,
    )
    assert isinstance(connection, _CloseFailureConnection)
    connection.row_factory = sqlite3.Row
    connection.close_timing = close_timing
    close_error = RuntimeError(f"capability close {close_timing} effect")
    connection.close_marker = close_error
    primary_error = MemoryError("connection capability construction failed")
    prior_error = LookupError("prior capability construction evidence")
    primary_error.__cause__ = prior_error
    primary_error.__suppress_context__ = True
    constructed: list[IndexStore] = []

    def open_instrumented_connection(_path: Path) -> sqlite3.Connection:
        return connection

    def fail_capability(provider: Any) -> None:
        constructed.append(cast(IndexStore, provider.__self__))
        raise primary_error

    with monkeypatch.context() as patch:
        patch.setattr(store_module, "_open_index_connection", open_instrumented_connection)
        patch.setattr(store_module, "_ConnectionCapability", fail_capability)
        with pytest.raises(MemoryError) as captured:
            IndexStore(database)

    assert captured.value is primary_error
    assert len(constructed) == 1
    assert constructed[0]._closed is True
    _assert_exception_identity_once(primary_error, prior_error)
    _assert_exception_identity_once(primary_error, close_error)
    _assert_acyclic_exception_graph(primary_error)
    with pytest.raises(sqlite3.ProgrammingError, match="closed"):
        sqlite3.Connection.execute(connection, "SELECT 1")
    _assert_writer_unblocked(database)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("rollback_timing", ["before", "after"])
@pytest.mark.parametrize("close_timing", ["before", "after"])
def test_schema_initialization_failure_preserves_primary_across_native_cleanup(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    journal_mode: str,
    rollback_timing: str,
    close_timing: str,
) -> None:
    import excel_lsp.core.index.store as store_module

    database = tmp_path / (
        f"schema-cleanup-{prefer_rtree}-{journal_mode}-{rollback_timing}-{close_timing}.db"
    )
    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_CloseFailureConnection,
    )
    assert isinstance(connection, _CloseFailureConnection)
    journal_row = connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()
    assert journal_row is not None
    assert str(journal_row[0]) == journal_mode
    connection.row_factory = sqlite3.Row
    connection.rollback_fails = True
    connection.rollback_timing = rollback_timing
    rollback_error = RuntimeError(f"schema rollback {rollback_timing} effect")
    connection.rollback_marker = rollback_error
    connection.close_timing = close_timing
    close_error = RuntimeError(f"schema close {close_timing} effect")
    connection.close_marker = close_error
    primary_error = ValueError("schema initialization locked body failed")
    prior_error = LookupError("prior schema initialization evidence")
    primary_error.__cause__ = prior_error
    primary_error.__suppress_context__ = True
    constructed: list[IndexStore] = []

    def open_instrumented_connection(_path: Path) -> sqlite3.Connection:
        return connection

    def fail_locked_initialization(self: IndexStore, *, prefer_rtree: bool) -> None:
        del prefer_rtree
        constructed.append(self)
        assert sqlite3.Connection.in_transaction.__get__(self._connection) is True
        raise primary_error

    with monkeypatch.context() as patch:
        patch.setattr(store_module, "_open_index_connection", open_instrumented_connection)
        patch.setattr(IndexStore, "_initialize_schema_locked", fail_locked_initialization)
        with pytest.raises(ValueError) as captured:
            IndexStore(database, prefer_rtree=prefer_rtree)

    assert captured.value is primary_error
    assert len(constructed) == 1
    assert constructed[0]._closed is True
    _assert_exception_identity_once(primary_error, prior_error)
    _assert_exception_identity_once(primary_error, rollback_error)
    _assert_exception_identity_once(primary_error, close_error)
    _assert_acyclic_exception_graph(primary_error)
    with pytest.raises(sqlite3.ProgrammingError, match="closed"):
        sqlite3.Connection.execute(connection, "SELECT 1")
    _assert_writer_unblocked(database)


@pytest.mark.parametrize("failure_point", ["tracker", "foreign-keys"])
@pytest.mark.parametrize("close_timing", ["before", "after"])
def test_open_connection_configuration_failure_preserves_primary_and_closes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    close_timing: str,
    failure_point: str,
) -> None:
    import excel_lsp.core.index.store as store_module

    database = tmp_path / f"configuration-cleanup-{failure_point}-{close_timing}.xlsp.db"
    original_connect = sqlite3.connect
    connection = original_connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_ConfigurationFailureConnection,
    )
    assert isinstance(connection, _ConfigurationFailureConnection)
    primary_error = RuntimeError(f"connection configuration failed at {failure_point}")
    prior_error = LookupError("prior connection configuration evidence")
    primary_error.__cause__ = prior_error
    primary_error.__suppress_context__ = True
    connection.configuration_marker = primary_error
    connection.configuration_failure_point = failure_point
    connection.close_timing = close_timing
    close_error = RuntimeError(f"configuration close {close_timing} effect")
    connection.close_marker = close_error

    def return_instrumented_connection(*args: Any, **kwargs: Any) -> sqlite3.Connection:
        del args, kwargs
        return connection

    with monkeypatch.context() as patch:
        patch.setattr(store_module.sqlite3, "connect", return_instrumented_connection)
        with pytest.raises(RuntimeError) as captured:
            store_module._open_index_connection(database)

    assert captured.value is primary_error
    _assert_exception_identity_once(primary_error, prior_error)
    _assert_exception_identity_once(primary_error, close_error)
    _assert_acyclic_exception_graph(primary_error)
    with pytest.raises(sqlite3.ProgrammingError, match="closed"):
        sqlite3.Connection.execute(connection, "SELECT 1")
    _assert_writer_unblocked(database)


@pytest.mark.parametrize("rollback_timing", ["before", "after"])
def test_native_rollback_fallback_finishes_before_connection_close(
    tmp_path: Path,
    rollback_timing: str,
) -> None:
    import excel_lsp.core.index.store as store_module

    database = tmp_path / f"native-rollback-{rollback_timing}.xlsp.db"
    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_CloseFailureConnection,
    )
    assert isinstance(connection, _CloseFailureConnection)
    connection.execute("BEGIN IMMEDIATE")
    rollback_error = RuntimeError(f"native rollback {rollback_timing} effect")
    connection.rollback_fails = True
    connection.rollback_timing = rollback_timing
    connection.rollback_marker = rollback_error
    cleanup_errors: list[BaseException] = []

    store_module._rollback_native_connection(connection, cleanup_errors)

    assert cleanup_errors == [rollback_error]
    descriptor = cast(Any, sqlite3.Connection.in_transaction)
    assert descriptor.__get__(connection, sqlite3.Connection) is False
    sqlite3.Connection.close(connection)
    _assert_writer_unblocked(database)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("close_timing", ["before", "after"])
def test_constructor_edge_store_failure_closes_before_assignment(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    journal_mode: str,
    close_timing: str,
) -> None:
    import excel_lsp.core.index.store as store_module

    database = tmp_path / (
        f"constructor-edge-store-{prefer_rtree}-{journal_mode}-{close_timing}.xlsp.db"
    )
    with IndexStore(database, prefer_rtree=prefer_rtree):
        pass

    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_CloseFailureConnection,
    )
    assert isinstance(connection, _CloseFailureConnection)
    journal_row = connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()
    assert journal_row is not None
    assert str(journal_row[0]) == journal_mode
    connection.row_factory = sqlite3.Row
    connection.close_timing = close_timing
    connection.close_marker = RuntimeError(f"edge-store close {close_timing} effect")
    connection.state_spoof = "false-positive" if close_timing == "before" else "false-negative"

    prior_error = LookupError("prior EdgeStore construction evidence")
    primary_error = RuntimeError("EdgeStore construction failed")
    primary_error.__cause__ = prior_error
    primary_error.__suppress_context__ = True
    constructed: list[IndexStore] = []
    initialize_schema = IndexStore._initialize_schema

    class FailingEdgeStore(EdgeStore):
        def __init__(self, connection: sqlite3.Connection) -> None:
            del connection
            raise primary_error

    def open_instrumented_connection(_path: Path) -> sqlite3.Connection:
        return connection

    def capture_initialization(self: IndexStore, *, prefer_rtree: bool) -> None:
        constructed.append(self)
        initialize_schema(self, prefer_rtree=prefer_rtree)

    with monkeypatch.context() as patch:
        patch.setattr(store_module, "_open_index_connection", open_instrumented_connection)
        patch.setattr(store_module, "EdgeStore", FailingEdgeStore)
        patch.setattr(IndexStore, "_initialize_schema", capture_initialization)
        with pytest.raises(RuntimeError) as captured:
            IndexStore(database, prefer_rtree=prefer_rtree)

    assert captured.value is primary_error
    assert len(constructed) == 1
    assert constructed[0]._closed is True
    assert not hasattr(constructed[0], "edge_store")
    assert connection.close_calls == 1
    assert connection.rollback_calls == 0
    _assert_exception_identity_once(primary_error, prior_error)
    _assert_exception_identity_once(primary_error, connection.close_marker)
    _assert_acyclic_exception_graph(primary_error)
    notes = getattr(primary_error, "__notes__", ())
    assert (
        notes.count(
            "IndexStore initialization cleanup also failed; the original error remains primary."
        )
        == 1
    )
    with pytest.raises(sqlite3.ProgrammingError, match="closed"):
        sqlite3.Connection.execute(connection, "SELECT 1")
    _assert_writer_unblocked(database)
    with IndexStore(database, prefer_rtree=prefer_rtree) as reopened:
        expected_backend = "rtree" if prefer_rtree else "interval"
        assert reopened.edge_store.backend == expected_backend


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_close_uses_rollback_fallback_finalizer_and_always_closes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
) -> None:
    store = _transaction_guard_store(tmp_path, prefer_rtree, "close-fallback")
    database = store.path
    before_export = store.canonical_export()
    before_generation = store.generation
    before_state = _graph_spatial_state(store)
    _start_abandoned_graph_transaction(store)

    def fail_primary_rollback() -> None:
        raise RuntimeError("injected close rollback failure")

    def fail_hook_before_restoration() -> None:
        raise RuntimeError("injected close hook failure")

    with monkeypatch.context() as patch:
        patch.setattr(store, "_rollback_after_failed_transaction", fail_primary_rollback)
        patch.setattr(
            store.edge_store,
            "transaction_rolled_back",
            fail_hook_before_restoration,
        )
        with pytest.raises(BaseExceptionGroup) as captured:
            store.close()

    messages = tuple(str(error) for error in captured.value.exceptions)
    assert messages == (
        "injected close rollback failure",
        "injected close hook failure",
    )
    with pytest.raises(RuntimeError, match="closed"):
        _ = store.connection
    store.close()

    writer = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.rollback()
    finally:
        writer.close()

    with IndexStore(database, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export
        assert reopened.generation == before_generation
        assert _graph_spatial_state(reopened) == before_state
        reopened.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
def test_close_authorizer_denial_still_releases_writer_and_marks_closed(
    tmp_path: Path,
    prefer_rtree: bool,
    journal_mode: str,
) -> None:
    store = _transaction_guard_store(
        tmp_path,
        prefer_rtree,
        f"close-authorizer-{journal_mode}",
    )
    database = store.path
    before_export = store.canonical_export()
    selected_mode = str(
        store.connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()[0]
    )
    assert selected_mode == journal_mode
    _start_abandoned_graph_transaction(store)

    def deny_rollback(
        action: int,
        argument_one: str | None,
        argument_two: str | None,
        database_name: str | None,
        trigger: str | None,
    ) -> int:
        del argument_two, database_name, trigger
        if action == sqlite3.SQLITE_TRANSACTION and argument_one == "ROLLBACK":
            return sqlite3.SQLITE_DENY
        return sqlite3.SQLITE_OK

    store.connection.set_authorizer(deny_rollback)
    with pytest.raises(BaseExceptionGroup):
        store.close()
    with pytest.raises(RuntimeError, match="closed"):
        _ = store.connection

    writer = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.rollback()
    finally:
        writer.close()

    with IndexStore(database, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export
        reopened.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("close_timing", ["before", "after"])
@pytest.mark.parametrize("rollback_mode", ["normal", "fallback"])
def test_close_retries_transient_physical_close_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    journal_mode: str,
    close_timing: str,
    rollback_mode: str,
) -> None:
    store = _transaction_guard_store(
        tmp_path,
        prefer_rtree,
        f"close-retry-{journal_mode}-{close_timing}-{rollback_mode}",
    )
    database = store.path
    before_export = store.canonical_export()
    selected_mode = str(
        store.connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()[0]
    )
    assert selected_mode == journal_mode
    _start_abandoned_graph_transaction(store)
    original_close = store._close_connection
    close_error = RuntimeError(f"close {close_timing} effect")

    def failing_close() -> None:
        if close_timing == "after":
            original_close()
        raise close_error

    def fail_primary_rollback() -> None:
        raise RuntimeError("injected primary rollback failure")

    with monkeypatch.context() as patch:
        patch.setattr(store, "_close_connection", failing_close)
        if rollback_mode == "fallback":
            patch.setattr(
                store,
                "_rollback_after_failed_transaction",
                fail_primary_rollback,
            )
        with pytest.raises((RuntimeError, ExceptionGroup)):
            store.close()

    with pytest.raises(RuntimeError, match="closed"):
        _ = store.connection
    store.close()

    writer = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.rollback()
    finally:
        writer.close()

    with IndexStore(database, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export
        reopened.edge_store.require_clean()


@pytest.mark.parametrize("connection_state", ["idle", "already-closed"])
def test_close_always_invokes_no_io_rollback_finalizer(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    connection_state: str,
) -> None:
    store = _transaction_guard_store(tmp_path, False, f"close-finalizer-{connection_state}")
    original_finalizer = store.edge_store.finalize_transaction_rollback
    finalizer_calls = 0

    def counted_finalizer() -> None:
        nonlocal finalizer_calls
        finalizer_calls += 1
        original_finalizer()

    monkeypatch.setattr(
        store.edge_store,
        "finalize_transaction_rollback",
        counted_finalizer,
    )
    if connection_state == "already-closed":
        store.connection.close()
        store.close()
    else:
        store.close()

    assert finalizer_calls == 1
    with pytest.raises(RuntimeError, match="closed"):
        _ = store.connection
    store.close()
    assert finalizer_calls == 1


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("body_chain", ["none", "cause", "context"])
def test_context_body_error_remains_primary_when_close_hook_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    body_chain: str,
) -> None:
    store = _transaction_guard_store(tmp_path, prefer_rtree, "close-body-error")
    database = store.path
    before_export = store.canonical_export()
    marker = ValueError("context body primary")
    prior_error = LookupError(f"prior body {body_chain}")

    def fail_hook_before_restoration() -> None:
        raise RuntimeError("injected context close hook failure")

    with monkeypatch.context() as patch:
        patch.setattr(
            store.edge_store,
            "transaction_rolled_back",
            fail_hook_before_restoration,
        )
        with pytest.raises(ValueError) as captured, store:
            _start_abandoned_graph_transaction(store)
            if body_chain == "cause":
                raise marker from prior_error
            if body_chain == "context":
                try:
                    raise prior_error
                except LookupError:
                    raise marker  # noqa: B904 - exercise implicit exception context
            raise marker

    assert captured.value is marker
    if body_chain == "none":
        assert isinstance(captured.value.__cause__, RuntimeError)
        assert str(captured.value.__cause__) == "injected context close hook failure"
    else:
        assert isinstance(captured.value.__cause__, BaseExceptionGroup)
        assert captured.value.__cause__.exceptions[0] is prior_error
        assert isinstance(captured.value.__cause__.exceptions[1], RuntimeError)
        assert str(captured.value.__cause__.exceptions[1]) == (
            "injected context close hook failure"
        )
    with pytest.raises(RuntimeError, match="closed"):
        _ = store.connection

    with IndexStore(database, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export
        reopened.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("begin_timing", ["before", "after"])
@pytest.mark.parametrize("release_mode", ["normal", "deny"])
def test_managed_begin_failure_is_released_or_conclusively_closed(
    tmp_path: Path,
    prefer_rtree: bool,
    journal_mode: str,
    begin_timing: str,
    release_mode: str,
) -> None:
    store, connection, before_export = _store_with_native_test_connection(
        tmp_path,
        prefer_rtree,
        journal_mode,
        f"begin-{begin_timing}-{release_mode}",
        _TransactionFailureConnection,
    )
    assert isinstance(connection, _TransactionFailureConnection)
    marker = sqlite3.DatabaseError(f"BEGIN IMMEDIATE {begin_timing} effect")
    connection.begin_timing = begin_timing
    connection.begin_marker = marker

    if release_mode == "deny":
        connection.set_authorizer(_deny_rollback)

    with pytest.raises(sqlite3.DatabaseError) as captured, store.transaction():
        raise AssertionError("managed transaction body must not be entered")
    assert captured.value is marker

    closed = begin_timing == "after" and release_mode == "deny"
    if closed:
        assert store._closed is True
        with pytest.raises(sqlite3.ProgrammingError):
            _ = connection.in_transaction
    else:
        assert store._closed is False
        assert connection.in_transaction is False
        connection.set_authorizer(None)
        with store.transaction():
            store.edge_store.require_clean()

    writer = sqlite3.connect(store.path, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.rollback()
    finally:
        writer.close()

    if not closed:
        store.close()
    with IndexStore(store.path, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export
        reopened.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("close_timing", ["before", "after"])
def test_close_uses_native_descriptor_after_persistent_virtual_failure(
    tmp_path: Path,
    prefer_rtree: bool,
    journal_mode: str,
    close_timing: str,
) -> None:
    store, connection, before_export = _store_with_native_test_connection(
        tmp_path,
        prefer_rtree,
        journal_mode,
        f"native-close-{close_timing}",
        _CloseFailureConnection,
    )
    assert isinstance(connection, _CloseFailureConnection)
    connection.close_timing = close_timing
    connection.close_marker = RuntimeError(f"native close {close_timing} effect")
    _start_abandoned_graph_transaction(store)
    connection.set_authorizer(_deny_rollback)

    with pytest.raises(BaseExceptionGroup):
        store.close()

    assert store._closed is True
    assert connection.close_calls == (1 if close_timing == "after" else 3)
    with pytest.raises(sqlite3.ProgrammingError):
        sqlite3.Connection.execute(connection, "SELECT 1")
    _assert_writer_unblocked(store.path)

    with IndexStore(store.path, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export
        reopened.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("close_timing", ["before", "after"])
@pytest.mark.parametrize("body_chain", ["success", "none", "cause", "context"])
def test_context_exit_conclusively_closes_native_subclass_and_keeps_body_primary(
    tmp_path: Path,
    prefer_rtree: bool,
    journal_mode: str,
    close_timing: str,
    body_chain: str,
) -> None:
    store, connection, before_export = _store_with_native_test_connection(
        tmp_path,
        prefer_rtree,
        journal_mode,
        f"native-exit-{close_timing}-{body_chain}",
        _CloseFailureConnection,
    )
    assert isinstance(connection, _CloseFailureConnection)
    connection.close_timing = close_timing
    connection.close_marker = RuntimeError(f"native context close {close_timing} effect")
    marker = ValueError("context body primary")
    prior_error = LookupError(f"prior native context body {body_chain}")
    _start_abandoned_graph_transaction(store)
    connection.set_authorizer(_deny_rollback)

    if body_chain != "success":
        with pytest.raises(ValueError) as captured, store:
            if body_chain == "cause":
                raise marker from prior_error
            if body_chain == "context":
                try:
                    raise prior_error
                except LookupError:
                    raise marker  # noqa: B904 - exercise implicit exception context
            raise marker
        assert captured.value is marker
        assert captured.value.__cause__ is not None
        if body_chain in {"cause", "context"}:
            assert isinstance(captured.value.__cause__, BaseExceptionGroup)
            assert captured.value.__cause__.exceptions[0] is prior_error
        _assert_acyclic_exception_graph(captured.value)
    else:
        with pytest.raises(BaseExceptionGroup), store:
            pass

    assert store._closed is True
    with pytest.raises(sqlite3.ProgrammingError):
        sqlite3.Connection.execute(connection, "SELECT 1")
    _assert_writer_unblocked(store.path)
    with IndexStore(store.path, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export
        reopened.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("body_chain", ["none", "cause", "context"])
def test_managed_transaction_cleanup_preserves_body_causal_evidence_without_cycles(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    body_chain: str,
) -> None:
    store = _transaction_guard_store(tmp_path, prefer_rtree, f"causal-{body_chain}")
    marker = ValueError("managed transaction body primary")
    prior_error = LookupError(f"prior managed body {body_chain}")
    rollback_error = RuntimeError("managed rollback cleanup failed")
    hook_error = RuntimeError("managed rollback bookkeeping failed")
    original_hook = store.edge_store.transaction_rolled_back

    def fail_rollback_before_effect() -> None:
        raise rollback_error

    def restore_bookkeeping_then_fail() -> None:
        original_hook()
        raise hook_error

    with monkeypatch.context() as patch:
        patch.setattr(store, "_rollback_after_failed_transaction", fail_rollback_before_effect)
        patch.setattr(store.edge_store, "transaction_rolled_back", restore_bookkeeping_then_fail)
        with pytest.raises(ValueError) as captured, store.transaction():
            if body_chain == "cause":
                raise marker from prior_error
            if body_chain == "context":
                try:
                    raise prior_error
                except LookupError:
                    raise marker  # noqa: B904 - exercise implicit exception context
            raise marker

    assert captured.value is marker
    cleanup_cause = captured.value.__cause__
    assert cleanup_cause is not None
    if body_chain != "none":
        assert isinstance(cleanup_cause, BaseExceptionGroup)
        assert cleanup_cause.exceptions[0] is prior_error
        cleanup_cause = cleanup_cause.exceptions[1]
    assert isinstance(cleanup_cause, BaseExceptionGroup)
    assert cleanup_cause.exceptions == (rollback_error, hook_error)
    _assert_acyclic_exception_graph(captured.value)
    assert store.connection.in_transaction is False
    store.edge_store.require_clean()
    store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("state_spoof", ["false-positive", "false-negative"])
@pytest.mark.parametrize("surface", ["close", "exit-success", "exit-error"])
def test_native_close_proof_bypasses_spoofed_in_transaction(
    tmp_path: Path,
    prefer_rtree: bool,
    state_spoof: str,
    surface: str,
) -> None:
    store, connection, before_export = _store_with_native_test_connection(
        tmp_path,
        prefer_rtree,
        "delete",
        f"state-spoof-{state_spoof}-{surface}",
        _CloseFailureConnection,
    )
    assert isinstance(connection, _CloseFailureConnection)
    _start_abandoned_graph_transaction(store)
    connection.set_authorizer(_deny_rollback)
    connection.close_timing = "before" if state_spoof == "false-positive" else "after"
    connection.close_marker = RuntimeError(f"spoof close {state_spoof}")
    connection.state_spoof = state_spoof
    assert sqlite3.Connection.execute(connection, "SELECT 1").fetchone()[0] == 1
    marker = ValueError("spoof context body primary")

    if surface == "close":
        with pytest.raises((RuntimeError, BaseExceptionGroup)):
            store.close()
    elif surface == "exit-success":
        with pytest.raises((RuntimeError, BaseExceptionGroup)), store:
            pass
    else:
        with pytest.raises(ValueError) as captured, store:
            raise marker
        assert captured.value is marker
        _assert_acyclic_exception_graph(captured.value)

    assert store._closed is True
    assert connection.close_calls == (3 if state_spoof == "false-positive" else 1)
    with pytest.raises(RuntimeError, match="closed"):
        _ = store.connection
    with pytest.raises(sqlite3.ProgrammingError):
        sqlite3.Connection.execute(connection, "SELECT 1")
    if state_spoof == "false-negative":
        with pytest.raises(RuntimeError, match="virtual state probe"):
            _ = connection.in_transaction
    _assert_writer_unblocked(store.path)
    with IndexStore(store.path, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export
        reopened.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("body_chain", ["cause", "context"])
@pytest.mark.parametrize("surface", ["transaction", "exit"])
def test_cleanup_composition_does_not_duplicate_prior_already_in_evidence(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    body_chain: str,
    surface: str,
) -> None:
    store = _transaction_guard_store(tmp_path, prefer_rtree, f"shared-{surface}-{body_chain}")
    marker = ValueError(f"{surface} body primary")
    prior_error = RuntimeError(f"shared {surface} {body_chain} evidence")
    second_cleanup_error = RuntimeError(f"second {surface} cleanup failure")

    def reuse_prior_as_cleanup_error() -> None:
        raise prior_error

    with monkeypatch.context() as patch:
        if surface == "transaction":
            patch.setattr(
                store,
                "_rollback_after_failed_transaction",
                reuse_prior_as_cleanup_error,
            )
            store.connection.set_authorizer(_deny_rollback)
            context = store.transaction()
        else:

            def fail_close_with_group() -> None:
                raise BaseExceptionGroup(
                    "shared close cleanup failures",
                    (prior_error, second_cleanup_error),
                )

            patch.setattr(store, "close", fail_close_with_group)
            context = store

        with pytest.raises(ValueError) as captured, context:
            if body_chain == "cause":
                raise marker from prior_error
            try:
                raise prior_error
            except RuntimeError:
                raise marker  # noqa: B904 - exercise implicit exception context

    assert captured.value is marker
    assert isinstance(captured.value.__cause__, BaseExceptionGroup)
    assert captured.value.__cause__.exceptions[0] is prior_error
    _assert_acyclic_exception_graph(captured.value)
    if surface == "transaction":
        assert store._closed is True
        store.close()
        _assert_writer_unblocked(store.path)
    else:
        assert store._closed is False
        store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("body_chain", ["cause", "context"])
@pytest.mark.parametrize("surface", ["transaction", "exit"])
def test_cleanup_boundaries_normalize_recursive_group_membership_aliases(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    body_chain: str,
    surface: str,
) -> None:
    store = _transaction_guard_store(tmp_path, prefer_rtree, f"nested-{surface}-{body_chain}")
    marker = ValueError(f"nested {surface} body primary")
    prior_error = RuntimeError(f"nested {surface} {body_chain} prior")
    cleanup_group, inner_group, distinct_one, distinct_two = _nested_store_cleanup_group(
        prior_error
    )

    def fail_with_nested_group() -> None:
        raise cleanup_group

    with monkeypatch.context() as patch:
        if surface == "transaction":
            patch.setattr(store, "_rollback_after_failed_transaction", fail_with_nested_group)
            context = store.transaction()
        else:
            patch.setattr(store, "_close_connection", fail_with_nested_group)
            context = store
        with pytest.raises(ValueError) as captured, context:
            if body_chain == "cause":
                raise marker from prior_error
            try:
                raise prior_error
            except RuntimeError:
                raise marker  # noqa: B904 - exercise implicit exception context

    assert captured.value is marker
    assert captured.value.__cause__ is cleanup_group
    assert cleanup_group.exceptions == (inner_group, distinct_two)
    assert inner_group.exceptions == (prior_error, distinct_one)
    assert distinct_two.__context__ is None
    _assert_acyclic_exception_graph(captured.value)
    if surface == "transaction":
        assert store.connection.in_transaction is False
        store.close()
    else:
        assert store._closed is True


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_direct_close_normalizes_later_nested_links_to_earlier_cleanup_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
) -> None:
    store = _transaction_guard_store(tmp_path, prefer_rtree, "direct-close-nested-link")
    _start_abandoned_graph_transaction(store)
    primary_error = RuntimeError("first direct-close cleanup failure")
    linked_error = RuntimeError("later direct-close member")
    linked_error.__context__ = primary_error
    linked_error.__suppress_context__ = True
    distinct_error = RuntimeError("distinct direct-close evidence")
    inner_group = BaseExceptionGroup("nested direct-close failure", (linked_error, distinct_error))
    cleanup_group = BaseExceptionGroup("outer direct-close failure", (inner_group,))

    def fail_first_rollback() -> None:
        raise primary_error

    def fail_finalizer_with_nested_group() -> None:
        raise cleanup_group

    with monkeypatch.context() as patch:
        patch.setattr(store, "_rollback_after_failed_transaction", fail_first_rollback)
        patch.setattr(
            store.edge_store,
            "finalize_transaction_rollback",
            fail_finalizer_with_nested_group,
        )
        with pytest.raises(BaseExceptionGroup) as captured:
            store.close()

    assert captured.value.exceptions == (primary_error, cleanup_group)
    assert inner_group.exceptions == (linked_error, distinct_error)
    assert linked_error.__context__ is None
    _assert_acyclic_exception_graph(captured.value)
    assert store._closed is True


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_post_commit_composition_detaches_nested_links_to_hook_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
) -> None:
    store = _transaction_guard_store(tmp_path, prefer_rtree, "commit-nested-link")
    hook_error = RuntimeError("commit hook primary")
    linked_error = RuntimeError("commit finalizer member")
    linked_error.__context__ = hook_error
    linked_error.__suppress_context__ = True
    distinct_error = RuntimeError("distinct commit finalizer evidence")
    inner_group = BaseExceptionGroup("nested commit finalizer", (linked_error, distinct_error))
    finalizer_group = BaseExceptionGroup("outer commit finalizer", (inner_group,))
    original_finalizer = store.edge_store.finalize_transaction_commit

    def publish_then_fail_hook() -> None:
        original_finalizer()
        raise hook_error

    def finalize_then_fail_group() -> None:
        original_finalizer()
        raise finalizer_group

    with monkeypatch.context() as patch:
        patch.setattr(store.edge_store, "transaction_committed", publish_then_fail_hook)
        patch.setattr(store.edge_store, "finalize_transaction_commit", finalize_then_fail_group)
        with pytest.raises(RuntimeError) as captured, store.transaction():
            store.set_meta("nested_commit_evidence", "durable")

    assert captured.value is hook_error
    assert captured.value.__cause__ is finalizer_group
    assert inner_group.exceptions == (linked_error, distinct_error)
    assert linked_error.__context__ is None
    _assert_acyclic_exception_graph(captured.value)
    assert store.get_meta("nested_commit_evidence") == "durable"
    store.edge_store.require_clean()
    store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_rank_key_catalog_is_primary_key_bounded_and_preserves_valid_duplicates(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    store = _transaction_guard_store(tmp_path, prefer_rtree, "rank-key-bounded")
    try:
        original = canonical_ranked_edges(store.connection)
        duplicate = RankedEdge(
            edge_id=2,
            source_sheet_id=original[0].source_sheet_id,
            source_rect=original[0].source_rect,
            destination_sheet_id=original[0].destination_sheet_id,
            destination_rect=original[0].destination_rect,
            dependent_rank=original[0].dependent_rank,
            precedent_rank=original[0].precedent_rank,
            dependent_key=original[0].dependent_key,
            precedent_key=original[0].precedent_key,
        )
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            )
            SELECT 2, src_kind, src_id, src_sheet_id, dst_sheet_id,
                   dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            FROM edges WHERE id = 1
            """
        )
        with store.transaction():
            store.edge_store.rebuild_ranked((*original, duplicate))

        plan = " ".join(
            str(row[3])
            for row in store.connection.execute(
                """
                EXPLAIN QUERY PLAN
                SELECT key_text FROM graph_rank_keys WHERE direction = ? AND rank = ?
                """,
                ("dependents", 1),
            )
        ).casefold()
        assert "primary key" in plan
        graph = store.dependency_graph
        result = graph.trace_dependents(GraphArea(2, "Destination", Rect(2, 2, 2, 2)))
        assert [child.target.symbol for child in result.root.children] == ["cell:Source!A1"]
    finally:
        store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("direction", ["dependents", "precedents"])
def test_active_mirror_mutation_invalidates_rank_identity_after_seal_restoration(
    tmp_path: Path,
    prefer_rtree: bool,
    direction: EdgeDirection,
) -> None:
    store = _transaction_guard_store(tmp_path, prefer_rtree, f"mirror-rank-{direction}")
    try:
        original = canonical_ranked_edges(store.connection)
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            )
            SELECT 2, src_kind, src_id, src_sheet_id, dst_sheet_id,
                   dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            FROM edges WHERE id = 1
            """
        )
        duplicate = RankedEdge(
            2,
            original[0].source_sheet_id,
            original[0].source_rect,
            original[0].destination_sheet_id,
            original[0].destination_rect,
            original[0].dependent_rank,
            original[0].precedent_rank,
            original[0].dependent_key,
            original[0].precedent_key,
        )
        with store.transaction():
            store.edge_store.rebuild_ranked((*original, duplicate))
        graph = store.dependency_graph
        if direction == "dependents":
            query = GraphArea(2, "Destination", Rect(2, 2, 2, 2))
            assert graph.trace_dependents(query).root.children
            table = store.edge_store.table_name
        else:
            query = GraphArea(1, "Source", Rect(1, 1, 1, 1))
            assert graph.trace_precedents(query).root.children
            table = store.edge_store.source_table_name
        sealed_state = tuple(
            store.connection.execute(
                """
                SELECT singleton, dirty, dependent_rank_max, precedent_rank_max,
                       revision, mutation_epoch, clean_epoch
                FROM graph_spatial_state
                """
            ).fetchone()
        )

        store.connection.execute(
            f"UPDATE {table} SET row_min = row_min + 1, row_max = row_max + 1 WHERE edge_id = 2"
        )
        store.connection.execute(
            """
            UPDATE graph_spatial_state
            SET singleton = ?, dirty = ?, dependent_rank_max = ?, precedent_rank_max = ?,
                revision = ?, mutation_epoch = ?, clean_epoch = ?
            """,
            sealed_state,
        )
        assert (
            store.connection.execute(
                "SELECT key_text FROM graph_rank_keys WHERE direction = ? AND rank = 1",
                (direction,),
            ).fetchone()
            is None
        )
        with pytest.raises(ExcelLSPError) as captured:
            if direction == "dependents":
                graph.trace_dependents(query)
            else:
                graph.trace_precedents(query)
        assert captured.value.code is ErrorCode.CORRUPT
    finally:
        store.close()


def test_edge_store_rtree_and_documented_interval_fallback_have_same_queries(
    tmp_path: Path,
) -> None:
    with IndexStore(tmp_path / "rtree.xlsp.db") as store:
        _exercise_edges(store.edge_store)

    fallback_connection = sqlite3.connect(":memory:", isolation_level=None)
    try:
        assert EdgeStore.ensure_schema(fallback_connection, prefer_rtree=False) == "interval"
        fallback = EdgeStore(fallback_connection)
        _exercise_edges(fallback)
        index_sql = fallback_connection.execute(
            "SELECT sql FROM sqlite_master WHERE name = 'edge_intervals_overlap'"
        ).fetchone()[0]
        assert "sheet_id, rank, row_min, row_max, col_min, col_max" in index_sql
    finally:
        fallback_connection.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_ranked_edge_mirrors_are_directional_bounded_and_dirty_gated(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    source = _descriptor("Source", order=0)
    destination = _descriptor("Destination", order=1)
    with IndexStore(
        tmp_path / f"ranked-{prefer_rtree}.xlsp.db", prefer_rtree=prefer_rtree
    ) as store:
        store.replace_sheet_catalog((source, destination))
        store.connection.executemany(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (?, 'cell', ?, 1, ?, ?, ?, ?, ?, ?)
            """,
            (
                (10, 65537, 2, 2, 2, 2, 2, "Ref"),
                (40, 65537, 2, 2, 2, 2, 2, "Ref"),
                (20, 131073, 2, 3, 3, 3, 3, "ref"),
                (30, 196609, None, None, None, None, None, "opaque:INDIRECT"),
            ),
        )
        with pytest.raises(RuntimeError, match="dirty"):
            store.edge_store.require_clean()

        ranked = (
            RankedEdge(10, 1, Rect(1, 1, 1, 1), 2, Rect(2, 2, 2, 2), 2, 1, ("dep-2",), ("pre-1",)),
            RankedEdge(40, 1, Rect(1, 1, 1, 1), 2, Rect(2, 2, 2, 2), 2, 1, ("dep-2",), ("pre-1",)),
            RankedEdge(20, 1, Rect(2, 2, 1, 1), 2, Rect(3, 3, 3, 3), 1, 2, ("dep-1",), ("pre-2",)),
            RankedEdge(30, 1, Rect(3, 3, 1, 1), None, None, 3, 3, ("dep-3",), ("pre-3",)),
        )
        with store.transaction():
            store.edge_store.rebuild_ranked(ranked)

        store.edge_store.require_clean()
        trust_state = store.connection.execute(
            """
            SELECT dirty, mutation_epoch, clean_epoch
            FROM graph_spatial_state WHERE singleton = 1
            """
        ).fetchone()
        assert trust_state is not None
        assert int(trust_state[0]) == 0
        assert int(trust_state[1]) == int(trust_state[2])
        assert store.edge_store.max_rank("dependents") == 3
        assert store.edge_store.max_rank("precedents") == 3
        destination_query = Rect(2, 3, 2, 3)
        assert (
            store.edge_store.first_matching_rank("dependents", 2, destination_query, after_rank=0)
            == 1
        )
        assert (
            store.edge_store.first_matching_rank("dependents", 2, destination_query, after_rank=1)
            == 2
        )
        assert store.edge_store.edge_id_at_rank("dependents", 2, destination_query, 1) == 20
        assert store.edge_store.edge_id_at_rank("dependents", 2, destination_query, 2) in {
            10,
            40,
        }
        assert store.edge_store.ranked_mirror(10, "dependents") == (
            2,
            Rect(2, 2, 2, 2),
            2,
        )
        assert store.edge_store.ranked_mirror(10, "precedents") == (
            1,
            Rect(1, 1, 1, 1),
            1,
        )
        assert (
            store.edge_store.first_matching_rank("precedents", 1, Rect(3, 3, 1, 1), after_rank=0)
            == 3
        )
        assert store.edge_store.ranked_mirror(30, "dependents") is None

        store.connection.execute(f"DELETE FROM {store.edge_store.table_name} WHERE edge_id = 10")
        with pytest.raises(RuntimeError, match="dirty"):
            store.edge_store.require_clean()
        with store.transaction():
            store.edge_store.rebuild_ranked(ranked)

        store.connection.execute("UPDATE edges SET via = 'changed' WHERE id = 10")
        with pytest.raises(RuntimeError, match="dirty"):
            store.edge_store.first_matching_rank("dependents", 2, destination_query)


def test_ranked_edge_rebuild_rejects_ranks_outside_rtree_i32(tmp_path: Path) -> None:
    with (
        IndexStore(tmp_path / "rank-overflow.xlsp.db") as store,
        pytest.raises(ValueError, match="32-bit"),
    ):
        store.edge_store.rebuild_ranked(
            (
                RankedEdge(
                    1,
                    1,
                    Rect(1, 1, 1, 1),
                    None,
                    None,
                    2_147_483_648,
                    1,
                    ("dependent",),
                    ("precedent",),
                ),
            )
        )


@pytest.mark.parametrize(
    ("second_dependent_rank", "second_dependent_key", "message"),
    (
        (1, ("different",), "rank maps to multiple semantic keys"),
        (2, ("same",), "semantic key maps to multiple ranks"),
    ),
)
def test_ranked_edge_rebuild_enforces_rank_key_bijection(
    tmp_path: Path,
    second_dependent_rank: int,
    second_dependent_key: tuple[object, ...],
    message: str,
) -> None:
    source = _descriptor("Source", order=0)
    with IndexStore(tmp_path / f"rank-key-{second_dependent_rank}.xlsp.db") as store:
        store.replace_sheet_catalog((source,))
        store.connection.executemany(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (?, 'cell', ?, 1, NULL, NULL, NULL, NULL, NULL, ?)
            """,
            ((1, 65537, "one"), (2, 131073, "two")),
        )
        first_key = ("same",) if second_dependent_rank == 2 else ("first",)
        with pytest.raises(ValueError, match=message):
            store.edge_store.rebuild_ranked(
                (
                    RankedEdge(
                        1,
                        1,
                        Rect(1, 1, 1, 1),
                        None,
                        None,
                        1,
                        1,
                        first_key,
                        ("precedent-1",),
                    ),
                    RankedEdge(
                        2,
                        1,
                        Rect(2, 2, 1, 1),
                        None,
                        None,
                        second_dependent_rank,
                        2,
                        second_dependent_key,
                        ("precedent-2",),
                    ),
                )
            )


def test_ranked_edge_rebuild_enforces_precedent_rank_key_bijection(tmp_path: Path) -> None:
    source = _descriptor("Source", order=0)
    with IndexStore(tmp_path / "precedent-rank-key.xlsp.db") as store:
        store.replace_sheet_catalog((source,))
        store.connection.executemany(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (?, 'cell', ?, 1, NULL, NULL, NULL, NULL, NULL, ?)
            """,
            ((1, 65537, "one"), (2, 131073, "two")),
        )
        invalid_cases = (
            (
                1,
                ("precedent-2",),
                "precedent rank maps to multiple semantic keys",
            ),
            (
                2,
                ("precedent-1",),
                "precedent semantic key maps to multiple ranks",
            ),
        )
        for second_rank, second_key, message in invalid_cases:
            with pytest.raises(ValueError, match=message):
                store.edge_store.rebuild_ranked(
                    (
                        RankedEdge(
                            1,
                            1,
                            Rect(1, 1, 1, 1),
                            None,
                            None,
                            1,
                            1,
                            ("dependent-1",),
                            ("precedent-1",),
                        ),
                        RankedEdge(
                            2,
                            1,
                            Rect(2, 2, 1, 1),
                            None,
                            None,
                            2,
                            second_rank,
                            ("dependent-2",),
                            second_key,
                        ),
                    )
                )


def test_public_graph_spatial_rebuild_bumps_generation_once_only_when_owning_transaction(
    tmp_path: Path,
) -> None:
    with IndexStore(tmp_path / "graph-generation.xlsp.db") as store:
        before = store.generation
        store.rebuild_graph_spatial_index()
        assert store.generation == before + 1

        with store.transaction():
            nested_before = store.generation
            store.rebuild_graph_spatial_index()
            assert store.generation == nested_before
        assert store.generation == nested_before


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("surface", ["index-store", "edge-store"])
def test_raw_transactions_cannot_adopt_graph_rebuilds_and_rollback_cleanly(
    tmp_path: Path, prefer_rtree: bool, surface: str
) -> None:
    with _transaction_guard_store(tmp_path, prefer_rtree, f"rollback-{surface}") as store:
        before_export = store.canonical_export()
        before_state = tuple(
            store.connection.execute(
                """
                SELECT dirty, dependent_rank_max, precedent_rank_max, revision,
                       mutation_epoch, clean_epoch
                FROM graph_spatial_state WHERE singleton = 1
                """
            ).fetchone()
        )
        before_mirror = store.edge_store.ranked_mirror(1, "dependents")
        before_rank = store.edge_store.first_matching_rank("dependents", 2, Rect(2, 2, 2, 2))
        ranked = canonical_ranked_edges(store.connection)
        consumed = False

        def tracked_records() -> Iterator[RankedEdge]:
            nonlocal consumed
            consumed = True
            yield from ranked

        store.connection.execute("BEGIN IMMEDIATE")
        store.connection.execute("UPDATE edges SET via = 'raw-change' WHERE id = 1")
        with pytest.raises(RuntimeError, match=r"with store\.transaction\(\)"):
            if surface == "index-store":
                store.rebuild_graph_spatial_index()
            else:
                store.edge_store.rebuild_ranked(tracked_records())
        if surface == "edge-store":
            assert consumed is False
        assert store.connection.in_transaction is True

        store.connection.rollback()
        assert store.canonical_export() == before_export
        after_state = tuple(
            store.connection.execute(
                """
                SELECT dirty, dependent_rank_max, precedent_rank_max, revision,
                       mutation_epoch, clean_epoch
                FROM graph_spatial_state WHERE singleton = 1
                """
            ).fetchone()
        )
        assert after_state == before_state
        # SQLite total_changes is monotonic even across rollback, so a raw
        # transaction permanently invalidates the cached facade. A fresh facade
        # revalidates the restored sidecar before serving it.
        with pytest.raises(RuntimeError, match="live seal"):
            store.edge_store.require_clean()
        fresh_edges = EdgeStore(store.connection)
        fresh_edges.require_clean()
        assert fresh_edges.ranked_mirror(1, "dependents") == before_mirror
        assert fresh_edges.first_matching_rank("dependents", 2, Rect(2, 2, 2, 2)) == before_rank


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_raw_transaction_commit_intent_is_refused_before_store_context_body(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _transaction_guard_store(tmp_path, prefer_rtree, "commit-intent") as store:
        before_export = store.canonical_export()
        entered = False

        store.connection.execute("BEGIN IMMEDIATE")
        with (
            pytest.raises(RuntimeError, match=r"with store\.transaction\(\)"),
            store.transaction(),
        ):
            entered = True
        assert entered is False
        assert store.connection.in_transaction is True

        # The raw transaction remains wholly caller-owned. Its caller may
        # explicitly commit it, but the store context never adopts that commit.
        store.connection.commit()
        assert store.connection.in_transaction is False
        assert store.canonical_export() == before_export
        store.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_nested_store_owned_graph_rebuild_remains_supported(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _transaction_guard_store(tmp_path, prefer_rtree, "nested") as store:
        generation = store.generation
        with store.transaction():
            store.connection.execute("UPDATE edges SET via = 'nested-change' WHERE id = 1")
            with store.transaction():
                store.rebuild_graph_spatial_index()
            store.edge_store.require_clean()

        assert store.generation == generation
        assert store.connection.execute("SELECT via FROM edges WHERE id = 1").fetchone()[0] == (
            "nested-change"
        )
        store.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_failed_deferred_fk_commit_restores_graph_transaction_state(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    with _transaction_guard_store(tmp_path, prefer_rtree, "failed-commit") as store:
        _create_deferred_fk_probe(store)
        query = GraphArea(2, "Destination", Rect(2, 2, 2, 2))
        before_export = store.canonical_export()
        before_generation = store.generation
        before_state = _graph_spatial_state(store)
        before_mirrors = (
            store.edge_store.ranked_mirror(1, "dependents"),
            store.edge_store.ranked_mirror(1, "precedents"),
        )
        before_query = store.dependency_graph.direct_dependents(query)

        with (
            pytest.raises(sqlite3.IntegrityError, match="FOREIGN KEY") as captured,
            store.transaction(),
        ):
            store.connection.execute("UPDATE edges SET via = 'uncommitted' WHERE id = 1")
            store.rebuild_graph_spatial_index()
            store.bump_generation()
            store.connection.execute("INSERT INTO deferred_fk_probe(id, sheet_id) VALUES (1, 9999)")
        assert captured.value.__cause__ is None

        assert store.connection.in_transaction is False
        assert store.connection.execute("SELECT COUNT(*) FROM deferred_fk_probe").fetchone()[0] == 0
        assert store.canonical_export() == before_export
        assert store.generation == before_generation
        assert _graph_spatial_state(store) == before_state
        assert (
            store.edge_store.ranked_mirror(1, "dependents"),
            store.edge_store.ranked_mirror(1, "precedents"),
        ) == before_mirrors
        store.edge_store.require_clean()
        assert store.dependency_graph.direct_dependents(query) == before_query

        with store.transaction():
            store.set_meta("post_failed_commit", "ok")
        assert store.get_meta("post_failed_commit") == "ok"


@pytest.mark.parametrize("cleanup_stage", ["rollback", "hook"])
def test_commit_cleanup_failure_is_chained_without_masking_commit_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, cleanup_stage: str
) -> None:
    with _transaction_guard_store(tmp_path, False, f"cleanup-{cleanup_stage}") as store:
        _create_deferred_fk_probe(store)
        original_hook = store.edge_store.transaction_rolled_back

        def fail_primary_rollback() -> None:
            raise RuntimeError("injected rollback cleanup failure")

        def restore_hook_then_fail() -> None:
            original_hook()
            raise RuntimeError("injected hook cleanup failure")

        with monkeypatch.context() as patch:
            if cleanup_stage == "rollback":
                patch.setattr(store, "_rollback_after_failed_transaction", fail_primary_rollback)
            else:
                patch.setattr(store.edge_store, "transaction_rolled_back", restore_hook_then_fail)

            with (
                pytest.raises(sqlite3.IntegrityError, match="FOREIGN KEY") as captured,
                store.transaction(),
            ):
                store.connection.execute("UPDATE edges SET via = 'uncommitted' WHERE id = 1")
                store.rebuild_graph_spatial_index()
                store.connection.execute(
                    "INSERT INTO deferred_fk_probe(id, sheet_id) VALUES (1, 9999)"
                )

        assert isinstance(captured.value.__cause__, RuntimeError)
        assert cleanup_stage in str(captured.value.__cause__)
        assert any("original error remains primary" in note for note in captured.value.__notes__)
        assert store.connection.in_transaction is False
        assert store.connection.execute("SELECT via FROM edges WHERE id = 1").fetchone()[0] == "ref"
        store.edge_store.require_clean()
        assert store.edge_store.first_matching_rank("dependents", 2, Rect(2, 2, 2, 2)) == 1


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("cleanup_stage", ["rollback", "hook"])
def test_body_failure_cleanup_is_chained_and_store_remains_safe(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    cleanup_stage: str,
) -> None:
    store = _transaction_guard_store(
        tmp_path,
        prefer_rtree,
        f"body-cleanup-{cleanup_stage}",
    )
    try:
        before_export = store.canonical_export()
        before_generation = store.generation
        before_state = _graph_spatial_state(store)
        query = GraphArea(2, "Destination", Rect(2, 2, 2, 2))
        before_query = store.dependency_graph.direct_dependents(query)
        original_hook = store.edge_store.transaction_rolled_back

        def fail_primary_rollback() -> None:
            raise RuntimeError("injected rollback cleanup failure")

        def restore_hook_then_fail() -> None:
            original_hook()
            raise RuntimeError("injected hook cleanup failure")

        with monkeypatch.context() as patch:
            if cleanup_stage == "rollback":
                patch.setattr(
                    store,
                    "_rollback_after_failed_transaction",
                    fail_primary_rollback,
                )
            else:
                patch.setattr(
                    store.edge_store,
                    "transaction_rolled_back",
                    restore_hook_then_fail,
                )

            with (
                pytest.raises(ValueError, match="primary body failure") as captured,
                store.transaction(),
            ):
                store.connection.execute("UPDATE edges SET via = 'uncommitted' WHERE id = 1")
                store.rebuild_graph_spatial_index()
                store.bump_generation()
                raise ValueError("primary body failure")

        assert isinstance(captured.value.__cause__, RuntimeError)
        assert cleanup_stage in str(captured.value.__cause__)
        assert any("original error remains primary" in note for note in captured.value.__notes__)
        assert store.connection.in_transaction is False
        assert store.canonical_export() == before_export
        assert store.generation == before_generation
        assert _graph_spatial_state(store) == before_state
        store.edge_store.require_clean()
        assert store.dependency_graph.direct_dependents(query) == before_query

        with store.transaction():
            store.set_meta("post_body_failure", "ok")
        assert store.get_meta("post_body_failure") == "ok"
    finally:
        store.close()

    with pytest.raises(RuntimeError, match="closed"):
        _ = store.connection


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("timing", ["before", "after"])
def test_rollback_hook_failure_is_backstopped_by_no_io_finalizer(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    timing: str,
) -> None:
    with _transaction_guard_store(
        tmp_path,
        prefer_rtree,
        f"rollback-hook-{timing}",
    ) as store:
        before_export = store.canonical_export()
        before_generation = store.generation
        before_state = _graph_spatial_state(store)
        query = GraphArea(2, "Destination", Rect(2, 2, 2, 2))
        before_query = store.dependency_graph.direct_dependents(query)
        original_hook = store.edge_store.transaction_rolled_back

        def failing_hook() -> None:
            if timing == "after":
                original_hook()
            raise RuntimeError(f"rollback hook {timing} restoration")

        with monkeypatch.context() as patch:
            patch.setattr(store.edge_store, "transaction_rolled_back", failing_hook)
            with (
                pytest.raises(ValueError, match="body failure") as captured,
                store.transaction(),
            ):
                store.connection.execute("UPDATE edges SET via = 'uncommitted' WHERE id = 1")
                store.rebuild_graph_spatial_index()
                store.bump_generation()
                raise ValueError("body failure")

        assert isinstance(captured.value.__cause__, RuntimeError)
        assert f"{timing} restoration" in str(captured.value.__cause__)
        assert store.connection.in_transaction is False
        assert store.canonical_export() == before_export
        assert store.generation == before_generation
        assert _graph_spatial_state(store) == before_state
        store.edge_store.require_clean()
        assert store.dependency_graph.direct_dependents(query) == before_query
        _assert_transaction_finalizer_is_idempotent_no_io(store, "rollback")

        with store.transaction():
            store.set_meta("post_rollback_hook_failure", timing)
        assert store.get_meta("post_rollback_hook_failure") == timing


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("timing", ["before", "after"])
def test_commit_hook_failure_preserves_durable_graph_and_never_rolls_back(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    timing: str,
) -> None:
    with _transaction_guard_store(
        tmp_path,
        prefer_rtree,
        f"commit-hook-{timing}",
    ) as store:
        original_hook = store.edge_store.transaction_committed
        rollback_calls = 0
        expected_export: dict[str, tuple[tuple[object, ...], ...]]
        expected_state: tuple[object, ...]
        expected_query: object
        expected_generation: int

        def failing_hook() -> None:
            if timing == "after":
                original_hook()
            raise RuntimeError(f"commit hook {timing} publication")

        def forbidden_rollback() -> None:
            nonlocal rollback_calls
            rollback_calls += 1
            raise AssertionError("rollback attempted after successful SQLite commit")

        query = GraphArea(2, "Destination", Rect(2, 2, 2, 2))
        with monkeypatch.context() as patch:
            patch.setattr(store.edge_store, "transaction_committed", failing_hook)
            patch.setattr(store, "_rollback_after_failed_transaction", forbidden_rollback)
            with (
                pytest.raises(RuntimeError, match=f"commit hook {timing}") as captured,
                store.transaction(),
            ):
                store.connection.execute("UPDATE edges SET via = 'committed' WHERE id = 1")
                store.rebuild_graph_spatial_index()
                expected_query = store.dependency_graph.direct_dependents(query)
                expected_generation = store.bump_generation()
                expected_export = store.canonical_export()
                expected_state = _graph_spatial_state(store)

        assert captured.value.__cause__ is None
        assert rollback_calls == 0
        assert store.connection.in_transaction is False
        assert store.canonical_export() == expected_export
        assert store.generation == expected_generation
        assert _graph_spatial_state(store) == expected_state
        store.edge_store.require_clean()
        assert store.dependency_graph.direct_dependents(query) == expected_query
        _assert_transaction_finalizer_is_idempotent_no_io(store, "commit")

        with store.transaction():
            store.set_meta("post_commit_hook_failure", timing)
        assert store.get_meta("post_commit_hook_failure") == timing


@pytest.mark.parametrize("phase", ["rollback", "commit"])
def test_transaction_finalizer_failure_is_chained_behind_primary_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    phase: str,
) -> None:
    with _transaction_guard_store(tmp_path, False, f"finalizer-{phase}") as store:
        if phase == "rollback":
            original_finalizer = store.edge_store.finalize_transaction_rollback

            def failing_hook() -> None:
                raise RuntimeError("rollback hook primary cleanup failure")

            def restoring_finalizer() -> None:
                original_finalizer()
                raise RuntimeError("rollback finalizer failure")

            with monkeypatch.context() as patch:
                patch.setattr(store.edge_store, "transaction_rolled_back", failing_hook)
                patch.setattr(
                    store.edge_store,
                    "finalize_transaction_rollback",
                    restoring_finalizer,
                )
                with (
                    pytest.raises(ValueError, match="body primary") as captured,
                    store.transaction(),
                ):
                    store.connection.execute("UPDATE edges SET via = 'uncommitted' WHERE id = 1")
                    store.rebuild_graph_spatial_index()
                    raise ValueError("body primary")

            assert isinstance(captured.value.__cause__, BaseExceptionGroup)
            messages = tuple(str(error) for error in captured.value.__cause__.exceptions)
            assert messages == (
                "rollback hook primary cleanup failure",
                "rollback finalizer failure",
            )
            assert store.connection.execute("SELECT via FROM edges WHERE id = 1").fetchone()[0] == (
                "ref"
            )
        else:
            original_finalizer = store.edge_store.finalize_transaction_commit

            def failing_hook() -> None:
                raise RuntimeError("commit hook primary")

            def publishing_finalizer() -> None:
                original_finalizer()
                raise RuntimeError("commit finalizer failure")

            with monkeypatch.context() as patch:
                patch.setattr(store.edge_store, "transaction_committed", failing_hook)
                patch.setattr(
                    store.edge_store,
                    "finalize_transaction_commit",
                    publishing_finalizer,
                )
                with (
                    pytest.raises(RuntimeError, match="commit hook primary") as captured,
                    store.transaction(),
                ):
                    store.connection.execute("UPDATE edges SET via = 'committed' WHERE id = 1")
                    store.rebuild_graph_spatial_index()

            assert isinstance(captured.value.__cause__, RuntimeError)
            assert str(captured.value.__cause__) == "commit finalizer failure"
            assert store.connection.execute("SELECT via FROM edges WHERE id = 1").fetchone()[0] == (
                "committed"
            )

        assert store.connection.in_transaction is False
        store.edge_store.require_clean()
        assert store.edge_store.first_matching_rank("dependents", 2, Rect(2, 2, 2, 2)) == 1
        with store.transaction():
            store.set_meta("post_finalizer_failure", phase)
        assert store.get_meta("post_finalizer_failure") == phase


@pytest.mark.parametrize("corruption", ["partial", "mixed"])
def test_current_schema_sidecar_with_partial_or_mixed_spatial_pair_is_rebuilt(
    tmp_path: Path, corruption: str
) -> None:
    database = tmp_path / f"spatial-{corruption}.xlsp.db"
    with IndexStore(database) as store:
        store.set_meta("generation", "7")

    connection = sqlite3.connect(database, isolation_level=None)
    try:
        if corruption == "partial":
            connection.execute("DROP TABLE edge_source_rtree")
        else:
            connection.execute(
                """
                CREATE TABLE edge_intervals(
                    edge_id INTEGER PRIMARY KEY, sheet_id INTEGER NOT NULL,
                    row_min INTEGER NOT NULL, row_max INTEGER NOT NULL,
                    col_min INTEGER NOT NULL, col_max INTEGER NOT NULL,
                    rank INTEGER NOT NULL
                )
                """
            )
        with pytest.raises(EdgeSchemaError, match=r"partial|mixes"):
            EdgeStore.ensure_schema(connection)
        with pytest.raises(EdgeSchemaError, match=r"partial|mixes"):
            EdgeStore(connection)
    finally:
        connection.close()

    with IndexStore(database) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == 8
        assert rebuilt.edge_store.backend == "rtree"
        assert {
            rebuilt.edge_store.table_name,
            rebuilt.edge_store.source_table_name,
        } == {"edge_rtree", "edge_source_rtree"}


@pytest.mark.parametrize(
    "corruption",
    (
        "missing-state-table",
        "missing-state-column",
        "malformed-state-column",
        "missing-rank-key-table",
        "missing-rank-key-trigger",
        "missing-base-trigger",
        "missing-mirror-trigger",
    ),
)
def test_current_schema_graph_state_or_required_trigger_damage_rebuilds_monotonically(
    tmp_path: Path, corruption: str
) -> None:
    database = tmp_path / f"graph-sidecar-{corruption}.xlsp.db"
    with IndexStore(database) as store:
        store.set_meta("generation", "17")

    connection = sqlite3.connect(database, isolation_level=None)
    try:
        if corruption == "missing-state-table":
            connection.execute("DROP TABLE graph_spatial_state")
        elif corruption == "missing-state-column":
            connection.execute("ALTER TABLE graph_spatial_state DROP COLUMN clean_epoch")
        elif corruption == "malformed-state-column":
            connection.execute("DROP TABLE graph_spatial_state")
            connection.execute(
                """
                CREATE TABLE graph_spatial_state (
                    singleton INTEGER PRIMARY KEY,
                    dirty INTEGER NOT NULL,
                    dependent_rank_max INTEGER NOT NULL,
                    precedent_rank_max INTEGER NOT NULL,
                    revision INTEGER NOT NULL,
                    mutation_epoch INTEGER NOT NULL,
                    clean_epoch TEXT NOT NULL
                )
                """
            )
            connection.execute("INSERT INTO graph_spatial_state VALUES (1, 1, 0, 0, 0, 0, '0')")
        elif corruption == "missing-rank-key-table":
            connection.execute("DROP TABLE graph_rank_keys")
        elif corruption == "missing-rank-key-trigger":
            connection.execute("DROP TRIGGER graph_rank_keys_graph_spatial_dirty_update")
        elif corruption == "missing-base-trigger":
            connection.execute("DROP TRIGGER edges_graph_spatial_dirty_update")
        else:
            connection.execute("DROP TRIGGER edge_rtree_rowid_graph_dirty_delete")

        with pytest.raises(
            EdgeSchemaError,
            match=r"graph spatial|graph dirty trigger|graph rank-key|graph_rank_keys",
        ):
            EdgeStore.ensure_schema(connection)
        with pytest.raises(
            EdgeSchemaError,
            match=r"graph spatial|graph dirty trigger|graph rank-key|graph_rank_keys",
        ):
            EdgeStore(connection)
    finally:
        connection.close()

    with IndexStore(database) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == 18
        columns = {
            str(row[1])
            for row in rebuilt.connection.execute("PRAGMA table_info(graph_spatial_state)")
        }
        assert {"mutation_epoch", "clean_epoch"} <= columns
        assert tuple(rebuilt.connection.execute("PRAGMA table_info(graph_rank_keys)"))
        trigger_count = int(
            rebuilt.connection.execute(
                """
                SELECT COUNT(*) FROM sqlite_master
                WHERE type = 'trigger' AND name LIKE '%graph%dirty%'
                """
            ).fetchone()[0]
        )
        assert trigger_count == 18


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    "bypass",
    ("comment-only", "when-zero", "undoing-body", "extra-undo-trigger"),
)
def test_current_schema_rejects_noncanonical_or_extra_graph_state_triggers(
    tmp_path: Path, prefer_rtree: bool, bypass: str
) -> None:
    database = tmp_path / f"trigger-{prefer_rtree}-{bypass}.xlsp.db"
    with IndexStore(database, prefer_rtree=prefer_rtree) as store:
        store.set_meta("generation", "23")
        physical_table = (
            "edge_rtree_rowid" if store.edge_store.backend == "rtree" else "edge_intervals"
        )

    connection = sqlite3.connect(database, isolation_level=None)
    try:
        trigger_name = f"{physical_table}_graph_dirty_insert"
        if bypass == "extra-undo-trigger":
            connection.execute(
                f"""
                CREATE TRIGGER extra_graph_state_undo
                AFTER INSERT ON {physical_table}
                BEGIN
                    UPDATE graph_spatial_state
                    SET dirty = 0, clean_epoch = mutation_epoch
                    WHERE singleton = 1;
                END
                """
            )
        else:
            connection.execute(f"DROP TRIGGER {trigger_name}")
            when_clause = "WHEN 0" if bypass == "when-zero" else ""
            comment = "/* comment-only mutation */" if bypass == "comment-only" else ""
            undo = (
                """
                UPDATE graph_spatial_state
                SET dirty = 0, clean_epoch = mutation_epoch
                WHERE singleton = 1;
                """
                if bypass == "undoing-body"
                else ""
            )
            connection.execute(
                f"""
                CREATE TRIGGER {trigger_name}
                AFTER INSERT ON {physical_table} {comment} {when_clause}
                BEGIN
                    UPDATE graph_spatial_state
                    SET dirty = 1, mutation_epoch = mutation_epoch + 1
                    WHERE singleton = 1;
                    {undo}
                END
                """
            )

        with pytest.raises(EdgeSchemaError, match=r"(?:unexpected.*trigger|trigger.*malformed)"):
            EdgeStore.ensure_schema(connection, prefer_rtree=prefer_rtree)
        with pytest.raises(EdgeSchemaError, match=r"(?:unexpected.*trigger|trigger.*malformed)"):
            EdgeStore(connection)
    finally:
        connection.close()

    with IndexStore(database, prefer_rtree=prefer_rtree) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == 24


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("rank_problem", ["null", "sparse"])
def test_open_validation_rejects_null_or_nondense_relational_graph_ranks(
    tmp_path: Path, prefer_rtree: bool, rank_problem: str
) -> None:
    source = _descriptor("Source", order=0)
    with IndexStore(
        tmp_path / f"rank-{prefer_rtree}-{rank_problem}.xlsp.db",
        prefer_rtree=prefer_rtree,
    ) as store:
        store.replace_sheet_catalog((source,))
        store.connection.executemany(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (?, 'cell', ?, 1, NULL, NULL, NULL, NULL, NULL, ?)
            """,
            ((1, 65537, "one"), (2, 131073, "two")),
        )
        store.edge_store.rebuild_ranked(
            (
                RankedEdge(
                    1,
                    1,
                    Rect(1, 1, 1, 1),
                    None,
                    None,
                    1,
                    1,
                    ("dependent-1",),
                    ("precedent-1",),
                ),
                RankedEdge(
                    2,
                    1,
                    Rect(2, 2, 1, 1),
                    None,
                    None,
                    2,
                    2,
                    ("dependent-2",),
                    ("precedent-2",),
                ),
            )
        )
        if rank_problem == "null":
            store.connection.execute("UPDATE edges SET dependent_rank = NULL WHERE id = 2")
        else:
            store.connection.execute(
                "UPDATE edges SET dependent_rank = 3, precedent_rank = 3 WHERE id = 2"
            )
            store.connection.execute(
                """
                UPDATE graph_spatial_state
                SET dependent_rank_max = 3, precedent_rank_max = 3
                WHERE singleton = 1
                """
            )
        store.connection.execute(
            """
            UPDATE graph_spatial_state
            SET dirty = 0, clean_epoch = mutation_epoch
            WHERE singleton = 1
            """
        )

        with pytest.raises(EdgeSchemaError, match="relational graph ranks"):
            EdgeStore(store.connection)


@pytest.mark.parametrize(
    ("prefer_rtree", "damage"),
    (
        (True, "missing-column"),
        (True, "wrong-virtual-module"),
        (True, "ordinary-table"),
        (False, "missing-column"),
        (False, "wrong-column-type"),
        (False, "wrong-table-identity"),
    ),
    ids=(
        "rtree-missing-column",
        "rtree-wrong-module",
        "rtree-ordinary-table",
        "interval-missing-column",
        "interval-wrong-type",
        "interval-wrong-identity",
    ),
)
def test_current_schema_rebuilds_malformed_physical_spatial_tables(
    tmp_path: Path, prefer_rtree: bool, damage: str
) -> None:
    database = tmp_path / f"physical-{prefer_rtree}-{damage}.xlsp.db"
    with IndexStore(database, prefer_rtree=prefer_rtree) as store:
        assert (store.edge_store.backend == "rtree") is prefer_rtree
        store.set_meta("generation", "31")

    connection = sqlite3.connect(database, isolation_level=None)
    try:
        if prefer_rtree:
            for operation in ("insert", "update", "delete"):
                connection.execute(f"DROP TRIGGER edge_rtree_rowid_graph_dirty_{operation}")
            connection.execute("DROP TABLE edge_rtree")
            if damage == "missing-column":
                connection.execute(
                    """
                    CREATE VIRTUAL TABLE edge_rtree USING rtree_i32(
                        edge_id, sheet_min, sheet_max, row_min, row_max, col_min, col_max
                    )
                    """
                )
            elif damage == "wrong-virtual-module":
                connection.execute(
                    """
                    CREATE VIRTUAL TABLE edge_rtree USING rtree(
                        edge_id, sheet_min, sheet_max, row_min, row_max,
                        col_min, col_max, rank_min, rank_max
                    )
                    """
                )
            else:
                connection.execute(
                    """
                    CREATE TABLE edge_rtree(
                        edge_id INT, sheet_min INT, sheet_max INT,
                        row_min INT, row_max INT, col_min INT, col_max INT,
                        rank_min INT, rank_max INT
                    )
                    """
                )
        elif damage == "missing-column":
            connection.execute("ALTER TABLE edge_intervals RENAME COLUMN rank TO rank_bad")
        else:
            for operation in ("insert", "update", "delete"):
                connection.execute(f"DROP TRIGGER edge_intervals_graph_dirty_{operation}")
            connection.execute("DROP TABLE edge_intervals")
            rank_definition = (
                "rank TEXT NOT NULL"
                if damage == "wrong-column-type"
                else ("rank INTEGER NOT NULL CHECK (rank >= 0)")
            )
            connection.execute(
                f"""
                CREATE TABLE edge_intervals(
                    edge_id INTEGER PRIMARY KEY,
                    sheet_id INTEGER NOT NULL,
                    row_min INTEGER NOT NULL,
                    row_max INTEGER NOT NULL,
                    col_min INTEGER NOT NULL,
                    col_max INTEGER NOT NULL,
                    {rank_definition}
                )
                """
            )

        with pytest.raises(EdgeSchemaError, match=r"physical|malformed"):
            EdgeStore.ensure_schema(connection, prefer_rtree=prefer_rtree)
        with pytest.raises(EdgeSchemaError, match=r"physical|malformed"):
            EdgeStore(connection)
    finally:
        connection.close()

    with IndexStore(database, prefer_rtree=prefer_rtree) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == 32


@pytest.mark.parametrize(
    "shadow_table",
    ["edge_rtree_node", "edge_source_rtree_rowid"],
    ids=["destination-node", "source-rowid"],
)
def test_corrupt_rtree_shadow_storage_recreates_database_monotonically(
    tmp_path: Path, shadow_table: str
) -> None:
    database = tmp_path / f"corrupt-{shadow_table}.xlsp.db"
    with IndexStore(database, prefer_rtree=True) as store:
        if store.edge_store.backend != "rtree":
            pytest.skip("SQLite RTree is unavailable")
        store.replace_sheet_catalog(
            (_descriptor("Source", order=0), _descriptor("Destination", order=1))
        )
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (1, 'cell', 65537, 1, 2, 2, 2, 2, 2, 'ref')
            """
        )
        store.rebuild_graph_spatial_index()
        store.set_meta("generation", "71")

    connection = sqlite3.connect(database, isolation_level=None)
    try:
        connection.execute(f"DELETE FROM {shadow_table}")
        with pytest.raises(EdgeSchemaError):
            EdgeStore.ensure_schema(connection, prefer_rtree=True)
        with pytest.raises(EdgeSchemaError):
            EdgeStore(connection)
    finally:
        connection.close()

    # A damaged RTree node can make valid DROP TABLE teardown itself fail.
    # Recovery closes the Windows handle and atomically starts a fresh index
    # file rather than using writable_schema to alter SQLite internals.
    with IndexStore(database, prefer_rtree=True) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.schema_created is False
        assert rebuilt.generation == 72
        assert rebuilt.edge_store.backend == "rtree"
        assert rebuilt.connection.execute("PRAGMA journal_mode").fetchone()[0] == "wal"
        assert rebuilt.connection.execute("PRAGMA foreign_keys").fetchone()[0] == 1
        assert rebuilt.connection.execute("SELECT COUNT(*) FROM edges").fetchone()[0] == 0
        assert rebuilt.connection.execute("PRAGMA integrity_check").fetchone()[0] == "ok"


def test_recreation_unlink_failure_is_not_masked_by_closed_connection_cleanup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    database = tmp_path / "corrupt-unlink-failure.xlsp.db"
    with IndexStore(database, prefer_rtree=True) as store:
        if store.edge_store.backend != "rtree":
            pytest.skip("SQLite RTree is unavailable")
        store.set_meta("generation", "91")

    connection = sqlite3.connect(database, isolation_level=None)
    try:
        connection.execute("DELETE FROM edge_rtree_node")
    finally:
        connection.close()

    target_wal = Path(f"{database}-wal")
    original_unlink = Path.unlink

    def fail_target_sidecar_unlink(path: Path, missing_ok: bool = False) -> None:
        if path == target_wal:
            raise OSError("simulated target sidecar unlink failure")
        original_unlink(path, missing_ok=missing_ok)

    with monkeypatch.context() as patch:
        patch.setattr(Path, "unlink", fail_target_sidecar_unlink)
        with pytest.raises(OSError, match="simulated target sidecar unlink failure"):
            IndexStore(database, prefer_rtree=True)
    assert database.exists()
    assert not tuple(tmp_path.glob("*.rebuild*"))

    # The failed atomic replacement leaves the original corrupt file in
    # place. A later open can retry and still preserves the generation rule.
    with IndexStore(database, prefer_rtree=True) as recovered:
        assert recovered.schema_rebuilt is True
        assert recovered.generation == 92


@pytest.mark.parametrize("close_timing", ["before", "after"])
def test_recreation_temp_descriptor_failure_preserves_primary_and_cleans_artifact(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    close_timing: str,
) -> None:
    import excel_lsp.core.index.store as store_module

    database = tmp_path / f"recreation-descriptor-{close_timing}.xlsp.db"
    store = IndexStore(database, prefer_rtree=False)
    store.set_meta("recreation_marker", "original")
    before_export = store.canonical_export()
    original_close = store_module.os.close
    close_error = OSError(f"temporary descriptor close {close_timing} effect")
    close_calls = 0

    def fail_first_descriptor_close(descriptor: int) -> None:
        nonlocal close_calls
        close_calls += 1
        if close_calls == 1:
            if close_timing == "after":
                original_close(descriptor)
            raise close_error
        original_close(descriptor)

    with monkeypatch.context() as patch:
        patch.setattr(store_module.os, "close", fail_first_descriptor_close)
        with pytest.raises(OSError) as captured:
            store._recreate_database(initial_generation=42, prefer_rtree=False)

    assert captured.value is close_error
    assert store._closed is True
    assert close_calls == 2
    assert not tuple(tmp_path.glob("*.rebuild*"))
    _assert_writer_unblocked(database)
    with IndexStore(database, prefer_rtree=False) as reopened:
        assert reopened.canonical_export() == before_export


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("close_timing", ["before", "after"])
def test_recreation_post_commit_close_failure_preserves_original_database(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    close_timing: str,
) -> None:
    import excel_lsp.core.index.store as store_module

    database = tmp_path / f"recreation-post-commit-{prefer_rtree}-{close_timing}.xlsp.db"
    store = IndexStore(database, prefer_rtree=prefer_rtree)
    store.set_meta("recreation_marker", "original")
    before_export = store.canonical_export()
    original_connect = sqlite3.connect
    replacements: list[_CloseFailureConnection] = []
    close_error = RuntimeError(f"post-commit replacement close {close_timing} effect")
    prior_error = LookupError("prior post-commit replacement close evidence")
    close_error.__cause__ = prior_error
    close_error.__suppress_context__ = True

    def open_replacement(path: Path) -> sqlite3.Connection:
        connection = original_connect(
            path,
            timeout=1.0,
            isolation_level=None,
            factory=_CloseFailureConnection,
        )
        assert isinstance(connection, _CloseFailureConnection)
        connection.row_factory = sqlite3.Row
        connection.close_timing = close_timing
        connection.close_marker = close_error
        replacements.append(connection)
        return connection

    with monkeypatch.context() as patch:
        patch.setattr(store_module, "_open_index_connection", open_replacement)
        with pytest.raises(RuntimeError) as captured:
            store._recreate_database(initial_generation=43, prefer_rtree=prefer_rtree)

    assert captured.value is close_error
    assert store._closed is True
    assert len(replacements) == 1
    replacement = replacements[0]
    _assert_exception_identity_once(close_error, prior_error)
    _assert_acyclic_exception_graph(close_error)
    with pytest.raises(sqlite3.ProgrammingError, match="closed"):
        sqlite3.Connection.execute(replacement, "SELECT 1")
    assert not tuple(tmp_path.glob("*.rebuild*"))
    _assert_writer_unblocked(database)
    with IndexStore(database, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
@pytest.mark.parametrize("rollback_timing", ["before", "after"])
@pytest.mark.parametrize("close_timing", ["before", "after"])
def test_recreation_build_failure_preserves_primary_and_releases_replacement(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    journal_mode: str,
    rollback_timing: str,
    close_timing: str,
) -> None:
    import excel_lsp.core.index.store as store_module

    database = tmp_path / (
        f"recreation-cleanup-{prefer_rtree}-{journal_mode}-{rollback_timing}-{close_timing}.xlsp.db"
    )
    store = IndexStore(database, prefer_rtree=prefer_rtree)
    store.set_meta("recreation_marker", "original")
    journal_row = store.connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()
    assert journal_row is not None
    assert str(journal_row[0]) == journal_mode
    before_export = store.canonical_export()
    original_connect = sqlite3.connect
    replacements: list[_CloseFailureConnection] = []
    primary_error = ValueError("replacement schema construction failed")
    prior_error = LookupError("prior replacement construction evidence")
    primary_error.__cause__ = prior_error
    primary_error.__suppress_context__ = True
    rollback_error = RuntimeError(f"replacement rollback {rollback_timing} effect")
    close_error = RuntimeError(f"replacement close {close_timing} effect")

    def open_replacement(path: Path) -> sqlite3.Connection:
        connection = original_connect(
            path,
            timeout=1.0,
            isolation_level=None,
            factory=_CloseFailureConnection,
        )
        assert isinstance(connection, _CloseFailureConnection)
        connection.row_factory = sqlite3.Row
        connection.rollback_fails = True
        connection.rollback_timing = rollback_timing
        connection.rollback_marker = rollback_error
        connection.close_timing = close_timing
        connection.close_marker = close_error
        connection.state_spoof = "false-positive" if close_timing == "before" else "false-negative"
        replacements.append(connection)
        return connection

    def fail_replacement_schema(
        self: IndexStore,
        *,
        initial_generation: int,
        prefer_rtree: bool,
    ) -> None:
        del self, initial_generation, prefer_rtree
        raise primary_error

    with monkeypatch.context() as patch:
        patch.setattr(store_module, "_open_index_connection", open_replacement)
        patch.setattr(IndexStore, "_create_schema", fail_replacement_schema)
        with pytest.raises(ValueError) as captured:
            store._recreate_database(initial_generation=41, prefer_rtree=prefer_rtree)

    assert captured.value is primary_error
    assert store._closed is True
    assert len(replacements) == 1
    replacement = replacements[0]
    _assert_exception_identity_once(primary_error, prior_error)
    _assert_exception_identity_once(primary_error, rollback_error)
    _assert_exception_identity_once(primary_error, close_error)
    _assert_acyclic_exception_graph(primary_error)
    with pytest.raises(sqlite3.ProgrammingError, match="closed"):
        sqlite3.Connection.execute(replacement, "SELECT 1")
    _assert_writer_unblocked(database)
    assert not tuple(tmp_path.glob("*.rebuild*"))
    with IndexStore(database, prefer_rtree=prefer_rtree) as reopened:
        assert reopened.canonical_export() == before_export


def test_healthy_interval_database_does_not_enter_file_recreation_recovery(
    tmp_path: Path,
) -> None:
    database = tmp_path / "healthy-interval.xlsp.db"
    with IndexStore(database, prefer_rtree=False) as store:
        assert store.edge_store.backend == "interval"
        store.set_meta("generation", "83")

    with IndexStore(database, prefer_rtree=False) as reopened:
        assert reopened.schema_rebuilt is False
        assert reopened.schema_created is False
        assert reopened.generation == 83
        assert reopened.edge_store.backend == "interval"


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_ranked_rebuild_seal_and_interval_cache_restore_after_outer_rollback(
    tmp_path: Path, prefer_rtree: bool
) -> None:
    source = _descriptor("Source", order=0)
    destination = _descriptor("Destination", order=1)
    with IndexStore(
        tmp_path / f"seal-rollback-{prefer_rtree}.xlsp.db",
        prefer_rtree=prefer_rtree,
    ) as store:
        store.replace_sheet_catalog((source, destination))
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (1, 'cell', 65537, 1, 2, 2, 2, 2, 2, 'original')
            """
        )
        ranked = (
            RankedEdge(
                1,
                1,
                Rect(1, 1, 1, 1),
                2,
                Rect(2, 2, 2, 2),
                1,
                1,
                ("dependent",),
                ("precedent",),
            ),
        )
        with store.transaction():
            store.edge_store.rebuild_ranked(ranked)
        store.edge_store.require_clean()
        original_epoch = int(
            store.connection.execute(
                "SELECT mutation_epoch FROM graph_spatial_state WHERE singleton = 1"
            ).fetchone()[0]
        )

        with (
            pytest.raises(RuntimeError, match="force rollback"),
            store.transaction(),
        ):
            store.connection.execute("UPDATE edges SET via = 'temporary' WHERE id = 1")
            store.edge_store.rebuild_ranked(ranked)
            store.edge_store.require_clean()
            raise RuntimeError("force rollback")

        store.edge_store.require_clean()
        assert store.connection.execute("SELECT via FROM edges WHERE id = 1").fetchone()[0] == (
            "original"
        )
        assert (
            int(
                store.connection.execute(
                    "SELECT mutation_epoch FROM graph_spatial_state WHERE singleton = 1"
                ).fetchone()[0]
            )
            == original_epoch
        )
        assert store.edge_store.first_matching_rank("dependents", 2, Rect(2, 2, 2, 2)) == 1


def test_interval_rank_bvh_is_work_bounded_over_50k_irrelevant_rectangles(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import excel_lsp.core.index.edges as edge_module

    source = _descriptor("Source", order=0)
    destination = _descriptor("Destination", order=1)
    visits_by_fanout: dict[int, int] = {}
    elapsed_by_fanout: dict[int, float] = {}
    with IndexStore(tmp_path / "interval-bvh.xlsp.db", prefer_rtree=False) as store:
        store.replace_sheet_catalog((source, destination))
        original_search = edge_module._interval_tree_min_rank
        visits = 0

        def counted_search(
            node: object,
            query: Rect,
            after_rank: int,
            best: int | None = None,
        ) -> int | None:
            nonlocal visits
            visits += 1
            return original_search(node, query, after_rank, best)  # type: ignore[arg-type]

        monkeypatch.setattr(edge_module, "_interval_tree_min_rank", counted_search)
        for fanout in (1_000, 10_000, 50_000):
            with store.transaction():
                store.connection.execute("DELETE FROM edges")
                relational_rows: list[tuple[object, ...]] = []
                ranked: list[RankedEdge] = []
                for index in range(fanout):
                    row = 1 if index == 0 else index + 100
                    edge_id = index + 1
                    rank = index + 1
                    packed = (row << 16) | 1
                    relational_rows.append((edge_id, packed, 1, 2, row, row, 1, 1, f"ref:{index}"))
                    ranked.append(
                        RankedEdge(
                            edge_id,
                            1,
                            Rect(row, row, 1, 1),
                            2,
                            Rect(row, row, 1, 1),
                            rank,
                            rank,
                            ("dependent", index),
                            ("precedent", index),
                        )
                    )
                store.connection.executemany(
                    """
                    INSERT INTO edges(
                        id, src_id, src_kind, src_sheet_id, dst_sheet_id,
                        dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                    ) VALUES (?, ?, 'cell', ?, ?, ?, ?, ?, ?, ?)
                    """,
                    relational_rows,
                )
                store.edge_store.rebuild_ranked(ranked)

            visits = 0
            started = perf_counter()
            for _iteration in range(200):
                assert store.edge_store.first_matching_rank("dependents", 2, Rect(1, 1, 1, 1)) == 1
                assert store.edge_store.first_matching_rank("precedents", 1, Rect(1, 1, 1, 1)) == 1
            elapsed_by_fanout[fanout] = perf_counter() - started
            visits_by_fanout[fanout] = visits // 200

    counts = tuple(visits_by_fanout.values())
    assert max(counts) <= min(counts) + 30, visits_by_fanout
    timings = tuple(elapsed_by_fanout.values())
    assert max(timings) <= min(timings) * 6 + 0.05, elapsed_by_fanout


def test_interval_rank_bvh_matches_brute_force_rectangles(tmp_path: Path) -> None:
    source = _descriptor("Source", order=0)
    destination = _descriptor("Destination", order=1)
    random = Random(20260727)
    with IndexStore(tmp_path / "interval-brute.xlsp.db", prefer_rtree=False) as store:
        store.replace_sheet_catalog((source, destination))
        relational_rows: list[tuple[object, ...]] = []
        ranked: list[RankedEdge] = []
        for index in range(1, 1_001):
            row_min = random.randint(1, 500)
            row_max = random.randint(row_min, min(550, row_min + 30))
            col_min = random.randint(1, 30)
            col_max = random.randint(col_min, min(40, col_min + 8))
            rect = Rect(row_min, row_max, col_min, col_max)
            relational_rows.append(
                (index, 65537, 1, 2, row_min, row_max, col_min, col_max, f"r:{index}")
            )
            ranked.append(
                RankedEdge(
                    index,
                    1,
                    rect,
                    2,
                    rect,
                    index,
                    index,
                    ("dependent", index),
                    ("precedent", index),
                )
            )
        with store.transaction():
            store.connection.executemany(
                """
                INSERT INTO edges(
                    id, src_id, src_kind, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES (?, ?, 'cell', ?, ?, ?, ?, ?, ?, ?)
                """,
                relational_rows,
            )
            store.edge_store.rebuild_ranked(ranked)

        for _query_index in range(250):
            row = random.randint(1, 550)
            col = random.randint(1, 40)
            after = random.randint(0, 1_000)
            query = Rect(row, row, col, col)
            expected = min(
                (
                    item.dependent_rank
                    for item in ranked
                    if item.dependent_rank > after
                    and item.destination_rect is not None
                    and item.destination_rect.intersects(query)
                ),
                default=None,
            )
            assert (
                store.edge_store.first_matching_rank("dependents", 2, query, after_rank=after)
                == expected
            )
            assert (
                store.edge_store.first_matching_rank("precedents", 1, query, after_rank=after)
                == expected
            )


def test_edges_allow_opaque_destinations_without_a_rectangle(tmp_path: Path) -> None:
    descriptor = _descriptor()
    with IndexStore(tmp_path / "opaque.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))
        store.connection.execute(
            """
            INSERT INTO edges(
                src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES ('cell', 65537, 1, NULL, NULL, NULL, NULL, NULL, 'opaque:INDIRECT')
            """
        )

        row = store.connection.execute(
            "SELECT dst_sheet_id, dst_row_min, dst_row_max, dst_col_min, dst_col_max FROM edges"
        ).fetchone()
        assert tuple(row) == (None, None, None, None, None)


def test_destination_sheet_reindex_preserves_incoming_edges(tmp_path: Path) -> None:
    source = _descriptor("Source", order=0)
    destination = _descriptor("Destination", order=1)
    with IndexStore(tmp_path / "incoming.xlsp.db") as store:
        store.replace_sheet_catalog((source, destination))
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (1, 'cell', 65537, 1, 2, 1, 1, 1, 1, 'ref')
            """
        )
        store.edge_store.insert(1, 2, Rect(1, 1, 1, 1))

        store.replace_sheet(
            destination,
            lambda _on_cell: SheetParseSummary(
                destination,
                "destination-new",
                1,
                1,
                0,
            ),
        )

        assert store.connection.execute("SELECT COUNT(*) FROM edges").fetchone()[0] == 1
        assert store.edge_store.query_point(2, 1, 1) == (1,)

        store.replace_sheet(
            source,
            lambda _on_cell: SheetParseSummary(source, "source-new", 1, 1, 0),
        )
        assert store.connection.execute("SELECT COUNT(*) FROM edges").fetchone()[0] == 0
        assert store.edge_store.query_point(2, 1, 1) == ()


def test_concurrent_cold_store_initialization_is_serialized(tmp_path: Path) -> None:
    workers = 8
    for run in range(5):
        database = tmp_path / f"concurrent-cold-{run}.xlsp.db"
        barrier = Barrier(workers)

        def open_store(
            _worker: int,
            start_barrier: Barrier = barrier,
            path: Path = database,
        ) -> tuple[str | None, int]:
            start_barrier.wait()
            with IndexStore(path) as store:
                return store.get_meta("schema_version"), store.generation

        with ThreadPoolExecutor(max_workers=workers) as pool:
            results = list(pool.map(open_store, range(workers)))

        assert results == [(SCHEMA_VERSION, 0)] * workers


def test_concurrent_schema_migration_runs_once(tmp_path: Path) -> None:
    database = tmp_path / "concurrent-migration.xlsp.db"
    with IndexStore(database) as store:
        store.set_meta("generation", 7)
        store.set_meta("schema_version", "obsolete")

    workers = 6
    barrier = Barrier(workers)

    def migrate(_worker: int) -> tuple[bool, str | None, int]:
        barrier.wait()
        with IndexStore(database) as store:
            return store.schema_rebuilt, store.get_meta("schema_version"), store.generation

    with ThreadPoolExecutor(max_workers=workers) as pool:
        results = list(pool.map(migrate, range(workers)))

    assert sum(rebuilt for rebuilt, _version, _generation in results) == 1
    assert {(version, generation) for _rebuilt, version, generation in results} == {
        (SCHEMA_VERSION, 8)
    }


def test_canonical_export_projects_surrogate_ids_and_uses_natural_order(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor()
    with IndexStore(tmp_path / "canonical.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))
        store.connection.execute(
            """
            INSERT INTO validations(
                id, sheet_id, row_min, row_max, col_min, col_max, allow_blank
            ) VALUES (91, 1, 5, 5, 2, 2, 0), (17, 1, 2, 2, 1, 1, 1)
            """
        )
        store.set_meta("generation", 44)
        store.set_meta("indexed_at", "different-on-every-build")

        exported = store.canonical_export()

        assert exported["sheets"] == (
            ("Data", "xl/worksheets/sheet1.xml", "", "worksheet", "visible", 0, 0),
        )
        assert exported["validations"] == (
            ("Data", 2, 2, 1, 1, None, None, None, None, 1),
            ("Data", 5, 5, 2, 2, None, None, None, None, 0),
        )
        assert all(key not in {"generation", "indexed_at"} for key, _value in exported["meta"])


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_canonical_spatial_export_preserves_edge_to_rectangle_association(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    source = _descriptor("Source", order=0)
    destination = _descriptor("Destination", order=1)
    with IndexStore(
        tmp_path / f"spatial-{prefer_rtree}.xlsp.db",
        prefer_rtree=prefer_rtree,
    ) as store:
        store.replace_sheet_catalog((source, destination))
        store.connection.executemany(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (?, 'cell', ?, 1, 2, ?, ?, ?, ?, 'ref')
            """,
            (
                (10, 65537, 1, 1, 1, 1),
                (20, 65538, 2, 2, 2, 2),
            ),
        )
        store.edge_store.insert(10, 2, Rect(1, 1, 1, 1))
        store.edge_store.insert(20, 2, Rect(2, 2, 2, 2))
        before = store.canonical_export()

        store.edge_store.insert(10, 2, Rect(2, 2, 2, 2))
        store.edge_store.insert(20, 2, Rect(1, 1, 1, 1))
        after = store.canonical_export()

        spatial_key = "edge_rtree" if prefer_rtree else "edge_intervals"
        assert before["edges"] == after["edges"]
        assert before[spatial_key] != after[spatial_key]


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_canonical_spatial_export_is_independent_of_physical_sheet_and_edge_ids(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    def build_export(
        database: Path,
        *,
        source_id: int,
        destination_id: int,
        edge_id: int,
    ) -> dict[str, tuple[tuple[object, ...], ...]]:
        with IndexStore(database, prefer_rtree=prefer_rtree) as store:
            expected_backend = "rtree" if prefer_rtree else "interval"
            assert store.edge_store.backend == expected_backend
            store.connection.executemany(
                """
                INSERT INTO sheets(
                    id, name, xml_part, part_hash, kind, visibility, max_row, max_col
                ) VALUES (?, ?, ?, ?, 'worksheet', 'visible', 2, 2)
                """,
                (
                    (source_id, "Source", "xl/worksheets/sheet1.xml", "source-hash"),
                    (
                        destination_id,
                        "Destination",
                        "xl/worksheets/sheet2.xml",
                        "destination-hash",
                    ),
                ),
            )
            store.connection.execute(
                """
                INSERT INTO edges(
                    id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES (?, 'cell', 65537, ?, ?, 2, 2, 2, 2, 'ref')
                """,
                (edge_id, source_id, destination_id),
            )
            store.edge_store.insert(edge_id, destination_id, Rect(2, 2, 2, 2))
            return store.canonical_export()

    first = build_export(
        tmp_path / f"canonical-low-ids-{prefer_rtree}.xlsp.db",
        source_id=1,
        destination_id=2,
        edge_id=10,
    )
    second = build_export(
        tmp_path / f"canonical-high-ids-{prefer_rtree}.xlsp.db",
        source_id=101,
        destination_id=307,
        edge_id=991,
    )

    assert first == second


def test_store_shapes_duplicate_catalog_and_cell_constraints_as_corrupt(
    tmp_path: Path,
) -> None:
    first = _descriptor("Data", order=0)
    duplicate_name = _descriptor("data", order=1)
    with IndexStore(tmp_path / "duplicate-catalog.xlsp.db") as store:
        with pytest.raises(ExcelLSPError) as catalog_error:
            store.replace_sheet_catalog((first, duplicate_name))
        assert catalog_error.value.code is ErrorCode.CORRUPT

    with IndexStore(tmp_path / "duplicate-cells.xlsp.db") as store:
        store.replace_sheet_catalog((first,))

        def parse(on_cell: object) -> SheetParseSummary:
            assert callable(on_cell)
            on_cell(CellRecord("A1", 1, 1, 1, "number"))
            on_cell(CellRecord("A1", 1, 1, 2, "number"))
            return SheetParseSummary(first, "duplicate", 1, 1, 2)

        with pytest.raises(ExcelLSPError) as cell_error:
            store.replace_sheet(first, parse)  # type: ignore[arg-type]
        assert cell_error.value.code is ErrorCode.CORRUPT
        assert store.connection.execute("SELECT COUNT(*) FROM cells").fetchone()[0] == 0


def test_large_sheet_warning_is_persisted_with_exact_related_json(tmp_path: Path) -> None:
    descriptor = _descriptor()
    with IndexStore(tmp_path / "large-sheet-warning.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))

        def parse(on_cell: object) -> SheetParseSummary:
            assert callable(on_cell)
            for row in range(1, 5):
                on_cell(CellRecord(f"A{row}", row, 1, row, "number"))
            return SheetParseSummary(descriptor, "large-sheet", 4, 1, 4)

        store.replace_sheet(
            descriptor,
            parse,  # type: ignore[arg-type]
            region_options=RegionOptions(
                large_sheet_threshold=3,
                large_dtype_sample_limit=1,
            ),
        )

        diagnostic = store.connection.execute(
            """
            SELECT severity, code, ref, related
            FROM diagnostics
            """
        ).fetchone()
        assert tuple(diagnostic) == (
            "warn",
            "W_LARGE_SHEET",
            "sheet:Data",
            '{"cellCount":4,"dtypeSampleLimit":1,"dtypeSampleStride":2}',
        )


def test_region_analysis_failure_rolls_back_replacement_and_generation(tmp_path: Path) -> None:
    descriptor = _descriptor()
    with IndexStore(tmp_path / "region-analysis-rollback.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))

        def initial_parse(on_cell: object) -> SheetParseSummary:
            assert callable(on_cell)
            on_cell(CellRecord("A1", 1, 1, "Header", "string"))
            on_cell(CellRecord("A2", 2, 1, 1, "number"))
            return SheetParseSummary(descriptor, "initial", 2, 1, 2)

        store.replace_sheet(descriptor, initial_parse)  # type: ignore[arg-type]
        before = store.canonical_export()
        generation = store.generation

        def corrupt_parse(on_cell: object) -> SheetParseSummary:
            assert callable(on_cell)
            on_cell(CellRecord("B1", 1, 2, "replacement", "string"))
            return SheetParseSummary(
                descriptor,
                "corrupt",
                1,
                2,
                1,
                tables=(TableInfo("BadTable", "BadTable", "not-a-range", 1, 0, ("Column",)),),
            )

        with pytest.raises(ExcelLSPError) as captured:
            store.replace_sheet(descriptor, corrupt_parse)  # type: ignore[arg-type]

        assert captured.value.code is ErrorCode.CORRUPT
        assert store.canonical_export() == before
        assert store.generation == generation


def _assert_transaction_finalizer_is_idempotent_no_io(store: IndexStore, finalizer: str) -> None:
    statements: list[str] = []
    store.connection.set_trace_callback(statements.append)
    try:
        if finalizer == "commit":
            store.edge_store.finalize_transaction_commit()
            store.edge_store.finalize_transaction_commit()
        else:
            assert finalizer == "rollback"
            store.edge_store.finalize_transaction_rollback()
            store.edge_store.finalize_transaction_rollback()
    finally:
        store.connection.set_trace_callback(None)
    assert statements == []


def _create_deferred_fk_probe(store: IndexStore) -> None:
    store.connection.execute(
        """
        CREATE TABLE deferred_fk_probe(
            id INTEGER PRIMARY KEY,
            sheet_id INTEGER NOT NULL,
            FOREIGN KEY(sheet_id) REFERENCES sheets(id)
                DEFERRABLE INITIALLY DEFERRED
        )
        """
    )
    # Test-only DDL intentionally advances SQLite's schema identity. Recreate
    # the facade after that setup so the transaction tests begin from a freshly
    # validated live seal.
    store.edge_store = EdgeStore(store.connection)


def _graph_spatial_state(store: IndexStore) -> tuple[object, ...]:
    row = store.connection.execute(
        """
        SELECT dirty, dependent_rank_max, precedent_rank_max, revision,
               mutation_epoch, clean_epoch
        FROM graph_spatial_state WHERE singleton = 1
        """
    ).fetchone()
    assert row is not None
    return tuple(row)


def _start_abandoned_graph_transaction(store: IndexStore) -> None:
    store.connection.execute("BEGIN IMMEDIATE")
    store.edge_store.transaction_started()
    store.connection.execute("UPDATE edges SET via = 'abandoned' WHERE id = 1")
    store.rebuild_graph_spatial_index()
    store.bump_generation()


class _TransactionFailureConnection(sqlite3.Connection):
    """Raise once immediately before or after a native BEGIN IMMEDIATE."""

    begin_timing: str | None = None
    begin_marker: sqlite3.DatabaseError

    def execute(self, sql: str, parameters: Any = (), /) -> sqlite3.Cursor:
        if self.begin_timing is not None and sql.strip().upper() == "BEGIN IMMEDIATE":
            timing = self.begin_timing
            self.begin_timing = None
            if timing == "after":
                super().execute(sql, parameters)
            raise self.begin_marker
        return super().execute(sql, parameters)


class _CloseFailureConnection(sqlite3.Connection):
    """Persistently fail virtual close before effect, or once after effect."""

    close_calls: int = 0
    close_timing: str = "before"
    close_marker: RuntimeError
    state_spoof: str | None = None
    rollback_calls: int = 0
    rollback_fails: bool = False
    rollback_timing: str = "before"
    rollback_marker: RuntimeError

    @property
    def in_transaction(self) -> bool:
        if self.state_spoof == "false-positive":
            raise sqlite3.ProgrammingError("Cannot operate on a closed database.")
        if self.state_spoof == "false-negative":
            raise RuntimeError("virtual state probe failed after native close")
        descriptor = cast(Any, sqlite3.Connection.in_transaction)
        return bool(descriptor.__get__(self, sqlite3.Connection))

    def close(self) -> None:
        self.close_calls += 1
        if self.close_timing == "after":
            super().close()
        raise self.close_marker

    def rollback(self) -> None:
        self.rollback_calls += 1
        if self.rollback_fails:
            if self.rollback_timing == "after":
                super().rollback()
            raise self.rollback_marker
        super().rollback()


class _ConfigurationFailureConnection(_CloseFailureConnection):
    """Fail connection tracking after sqlite3.connect returns a live handle."""

    configuration_marker: BaseException
    configuration_failure_point: str

    def install_graph_tracker(self) -> None:
        if self.configuration_failure_point == "tracker":
            raise self.configuration_marker

    def execute(self, sql: str, parameters: Any = (), /) -> sqlite3.Cursor:
        if (
            self.configuration_failure_point == "foreign-keys"
            and sql.strip().upper() == "PRAGMA FOREIGN_KEYS = ON"
        ):
            raise self.configuration_marker
        return super().execute(sql, parameters)


def _deny_rollback(
    action: int,
    argument_one: str | None,
    argument_two: str | None,
    database_name: str | None,
    trigger: str | None,
) -> int:
    del argument_two, database_name, trigger
    if action == sqlite3.SQLITE_TRANSACTION and argument_one == "ROLLBACK":
        return sqlite3.SQLITE_DENY
    return sqlite3.SQLITE_OK


def _store_with_native_test_connection(
    tmp_path: Path,
    prefer_rtree: bool,
    journal_mode: str,
    suffix: str,
    factory: type[sqlite3.Connection],
) -> tuple[IndexStore, sqlite3.Connection, dict[str, tuple[tuple[object, ...], ...]]]:
    store = _transaction_guard_store(tmp_path, prefer_rtree, suffix)
    before_export = store.canonical_export()
    database = store.path
    store.close()
    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=factory,
    )
    assert str(connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()[0]) == (
        journal_mode
    )
    store._connection = connection
    store._closed = False
    store.edge_store = EdgeStore(connection)
    return store, connection, before_export


def _assert_writer_unblocked(database: Path) -> None:
    writer = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    try:
        writer.execute("BEGIN IMMEDIATE")
        writer.rollback()
    finally:
        writer.close()


def _assert_acyclic_exception_graph(root: BaseException) -> None:
    seen: set[int] = set()
    pending = [root]
    while pending:
        error = pending.pop()
        assert id(error) not in seen
        seen.add(id(error))
        if error.__cause__ is not None:
            pending.append(error.__cause__)
        if error.__context__ is not None:
            pending.append(error.__context__)
        if isinstance(error, BaseExceptionGroup):
            pending.extend(error.exceptions)


def _assert_exception_identity_once(root: BaseException, target: BaseException) -> None:
    assert _exception_identity_count(root, target) == 1


def _assert_exception_identity_absent(root: BaseException, target: BaseException) -> None:
    assert _exception_identity_count(root, target) == 0


def _exception_identity_count(root: BaseException, target: BaseException) -> int:
    count = 0
    expanded: set[int] = set()
    pending = [root]
    while pending:
        error = pending.pop()
        if error is target:
            count += 1
        if id(error) in expanded:
            continue
        expanded.add(id(error))
        if error.__cause__ is not None:
            pending.append(error.__cause__)
        if error.__context__ is not None:
            pending.append(error.__context__)
        if isinstance(error, BaseExceptionGroup):
            pending.extend(error.exceptions)
    return count


def _nested_store_cleanup_group(
    prior_error: BaseException,
) -> tuple[BaseExceptionGroup, BaseExceptionGroup, RuntimeError, RuntimeError]:
    distinct_one = RuntimeError("distinct nested store cleanup one")
    distinct_two = RuntimeError("distinct nested store cleanup two")
    distinct_two.__context__ = prior_error
    distinct_two.__suppress_context__ = True
    inner_group = BaseExceptionGroup(
        "inner nested store cleanup",
        (prior_error, distinct_one),
    )
    outer_group = BaseExceptionGroup(
        "outer nested store cleanup",
        (inner_group, distinct_two),
    )
    return outer_group, inner_group, distinct_one, distinct_two


def _transaction_guard_store(tmp_path: Path, prefer_rtree: bool, suffix: str) -> IndexStore:
    store = IndexStore(
        tmp_path / f"transaction-guard-{prefer_rtree}-{suffix}.xlsp.db",
        prefer_rtree=prefer_rtree,
    )
    try:
        store.replace_sheet_catalog(
            (_descriptor("Source", order=0), _descriptor("Destination", order=1))
        )
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (1, 'cell', 65537, 1, 2, 2, 2, 2, 2, 'ref')
            """
        )
        store.rebuild_graph_spatial_index()
    except BaseException:
        store.close()
        raise
    return store


def _exercise_edges(edges: EdgeStore) -> None:
    edges.insert(9, 1, Rect(1, 10, 2, 2))
    edges.insert(4, 1, Rect(20, 30, 2, 4))
    edges.insert(7, 2, Rect(1, 100, 1, 10))
    assert edges.query_point(1, 5, 2) == (9,)
    assert edges.query_point(1, 15, 2) == ()
    assert edges.query_range(1, Rect(8, 22, 1, 3)) == (4, 9)
    assert edges.query_range(2, Rect(50, 50, 5, 5)) == (7,)
    assert edges.query_range_page(1, Rect(1, 30, 1, 4), limit=1) == (4,)
    assert edges.query_range_page(1, Rect(1, 30, 1, 4), after_edge_id=4, limit=1) == (9,)
    assert tuple(edges.iter_query_range(1, Rect(1, 30, 1, 4), page_size=1)) == (4, 9)
    assert tuple(edges.iter_query_point(2, 50, 5, page_size=1)) == (7,)
    assert edges.destination(7) == (2, Rect(1, 100, 1, 10))
    assert edges.destination(8) is None
    with pytest.raises(ValueError, match="limit"):
        edges.query_range_page(1, Rect(1, 1, 1, 1), limit=0)
    with pytest.raises(ValueError, match="after_edge_id"):
        edges.query_point_page(1, 1, 1, after_edge_id=-1)
    with pytest.raises(ValueError, match="limit"):
        edges.query_point_page(1, 1, 1, limit=EdgeStore.MAX_PAGE_SIZE + 1)
    edges.delete(9)
    assert edges.query_point(1, 5, 2) == ()
