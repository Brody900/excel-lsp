"""Persistent per-workbook SQLite index storage."""

from __future__ import annotations

import json
import math
import os
import sqlite3
import tempfile
import time
from collections.abc import Callable, Generator, Iterator, Mapping, Sequence
from contextlib import contextmanager, suppress
from dataclasses import dataclass
from pathlib import Path
from types import TracebackType
from typing import TYPE_CHECKING, Any, Self, cast

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.exception_evidence import (
    normalize_exception_graph,
    prepare_chained_failure_with_primary_evidence,
)
from excel_lsp.core.formulas.a1 import CellRef, resolve_reference
from excel_lsp.core.formulas.analysis import analyze_formula
from excel_lsp.core.formulas.blocks import FormulaBlock, FormulaCell
from excel_lsp.core.formulas.indexing import (
    SheetFormulaAnalysis,
    analyze_sheet_formulas,
)
from excel_lsp.core.formulas.references import (
    FormulaAnchor,
    ReferenceContext,
    TableBinding,
)
from excel_lsp.core.formulas.translation import translate_a1_formula
from excel_lsp.core.index.edges import (
    EdgeSchemaError,
    EdgeStore,
    GraphProjectionError,
    SQLiteConnectionLike,
    canonical_ranked_edges,
)
from excel_lsp.core.index.schema import (
    BASE_SCHEMA_SQL,
    CONTENT_TABLES_DELETE_ORDER,
    SCHEMA_VERSION,
)
from excel_lsp.core.models import (
    CellRecord,
    CellScalar,
    CellValueType,
    DataTableFormulaInfo,
    Rect,
    SheetDescriptor,
    SheetParseSummary,
    TableInfo,
    WorkbookMetadata,
)
from excel_lsp.core.parse.coordinates import make_cell_ref, parse_rect
from excel_lsp.core.parse.styles import DEFAULT_STYLE_CATALOG, StyleCatalog
from excel_lsp.core.regions import (
    RegionAnalysis,
    RegionCell,
    RegionOptions,
    analyze_sheet_regions,
)
from excel_lsp.core.symbols import cell_symbol_id, formula_block_symbol_id
from excel_lsp.core.values import JsonScalar, normalize_value

if TYPE_CHECKING:
    from excel_lsp.core.formulas.references import ExtractedReference
    from excel_lsp.core.graph.circular import BlockKey, CellNode, CircularBlock
    from excel_lsp.core.graph.queries import DependencyGraph

CellConsumer = Callable[[CellRecord], None]
SheetParser = Callable[[CellConsumer], SheetParseSummary]

_CELL_INSERT_SQL = """
INSERT INTO cells(
    sheet_id, row, col, ref, value, value_type, formula, style_idx,
    formula_kind, shared_index, array_ref, data_table
) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

_SQLITE_INT_MIN = -(1 << 63)
_SQLITE_INT_MAX = (1 << 63) - 1
_INITIALIZATION_RETRY_DELAYS = (0.01, 0.025, 0.05, 0.1, 0.2, 0.4)
_P3_DIAGNOSTIC_CODES = (
    "I_DYNAMIC_REF",
    "W_INCONSISTENT_FORMULA",
    "W_PARSE",
    "W_UNKNOWN_NAME",
)
_P4_DIAGNOSTIC_CODES = ("E_CIRCULAR", "W_POSSIBLE_CIRCULAR")
_PACKED_CELL_FACTOR = 1 << 16
_PSEUDO_CELL_BLOCK_BASE = 1 << 48


@dataclass(frozen=True, slots=True)
class _CircularOwnerNode:
    """Balanced rectangle index for exact formula-cell owner lookup."""

    bounds: Rect
    entries: tuple[CircularBlock, ...] = ()
    left: _CircularOwnerNode | None = None
    right: _CircularOwnerNode | None = None


class _DatabaseRecreationRequired(RuntimeError):
    """Signal that corrupt storage prevented transactional object teardown."""

    def __init__(self, initial_generation: int) -> None:
        super().__init__("the index database must be recreated")
        self.initial_generation = initial_generation


_GRAPH_AUTHORITATIVE_TABLES = frozenset(
    {
        "edges",
        "fblocks",
        "sheets",
        "graph_rank_keys",
        "graph_spatial_state",
        "edge_intervals",
        "edge_source_intervals",
        "sqlite_master",
        "sqlite_schema",
        "sqlite_temp_master",
        "sqlite_temp_schema",
    }
)
_SCHEMA_AUTHORIZE_ACTIONS = frozenset(
    getattr(sqlite3, name)
    for name in (
        "SQLITE_ALTER_TABLE",
        "SQLITE_CREATE_INDEX",
        "SQLITE_CREATE_TABLE",
        "SQLITE_CREATE_TEMP_INDEX",
        "SQLITE_CREATE_TEMP_TABLE",
        "SQLITE_CREATE_TEMP_TRIGGER",
        "SQLITE_CREATE_TRIGGER",
        "SQLITE_CREATE_VTABLE",
        "SQLITE_DROP_INDEX",
        "SQLITE_DROP_TABLE",
        "SQLITE_DROP_TEMP_INDEX",
        "SQLITE_DROP_TEMP_TABLE",
        "SQLITE_DROP_TEMP_TRIGGER",
        "SQLITE_DROP_TRIGGER",
        "SQLITE_DROP_VTABLE",
    )
    if hasattr(sqlite3, name)
)


def _is_graph_authorizer_action(
    action: int,
    argument_one: str | None,
    argument_two: str | None,
) -> bool:
    """Return whether one allowed statement can change graph trust identity."""
    if action in _SCHEMA_AUTHORIZE_ACTIONS:
        return True
    if action == sqlite3.SQLITE_PRAGMA:
        return argument_two is not None and (argument_one or "").casefold() in {
            "schema_version",
            "writable_schema",
        }
    if action not in {sqlite3.SQLITE_INSERT, sqlite3.SQLITE_UPDATE, sqlite3.SQLITE_DELETE}:
        return False
    table = (argument_one or argument_two or "").casefold()
    return table in _GRAPH_AUTHORITATIVE_TABLES or table.startswith(
        ("edge_rtree", "edge_source_rtree")
    )


class _TrackedConnection(sqlite3.Connection):
    """Private SQLite handle with a chained graph-write identity."""

    _graph_write_epoch: int
    _client_authorizer: Callable[[int, str | None, str | None, str | None, str | None], int] | None

    def install_graph_tracker(self) -> None:
        """Install tracking immediately after sqlite3.connect constructs us."""
        self._graph_write_epoch = 0
        self._client_authorizer = None
        sqlite3.Connection.set_authorizer(self, self._dispatch_authorizer)

    def set_authorizer(
        self,
        authorizer_callback: Callable[[int, str | None, str | None, str | None, str | None], int]
        | None,
    ) -> None:
        """Chain client policy without allowing it to displace graph tracking."""
        self._client_authorizer = authorizer_callback
        sqlite3.Connection.set_authorizer(self, self._dispatch_authorizer)

    def _dispatch_authorizer(
        self,
        action: int,
        argument_one: str | None,
        argument_two: str | None,
        database_name: str | None,
        trigger_name: str | None,
    ) -> int:
        graph_action = _is_graph_authorizer_action(action, argument_one, argument_two)
        if self._client_authorizer is not None:
            verdict = self._client_authorizer(
                action,
                argument_one,
                argument_two,
                database_name,
                trigger_name,
            )
            if verdict != sqlite3.SQLITE_OK:
                # SQLITE_DENY aborts before mutation. SQLITE_IGNORE can instead
                # let a write continue with altered semantics (notably DELETE),
                # so conservatively invalidate graph authority before returning.
                if verdict == sqlite3.SQLITE_IGNORE and graph_action:
                    self._graph_write_epoch += 1
                return verdict
        if graph_action:
            self._graph_write_epoch += 1
        return sqlite3.SQLITE_OK


class _CursorCapability:
    """Narrow cursor facade that never exposes its native connection."""

    __slots__ = ("__connection", "__cursor")

    def __init__(self, connection: _ConnectionCapability, cursor: sqlite3.Cursor) -> None:
        self.__connection = connection
        self.__cursor = cursor

    @property
    def connection(self) -> _ConnectionCapability:
        return self.__connection

    @property
    def description(self) -> tuple[tuple[Any, ...], ...] | None:
        return self.__cursor.description

    @property
    def lastrowid(self) -> int | None:
        return self.__cursor.lastrowid

    @property
    def rowcount(self) -> int:
        return self.__cursor.rowcount

    @property
    def arraysize(self) -> int:
        return self.__cursor.arraysize

    @arraysize.setter
    def arraysize(self, value: int) -> None:
        self.__cursor.arraysize = value

    @property
    def row_factory(self) -> None:
        """Public cursors intentionally cannot install native-cursor callbacks."""
        return None

    @row_factory.setter
    def row_factory(self, value: object) -> None:
        del value
        raise TypeError("custom cursor row factories are not supported")

    def execute(self, sql: str, parameters: Any = ()) -> Self:
        self.__cursor.execute(sql, parameters)
        return self

    def executemany(self, sql: str, seq_of_parameters: Any) -> Self:
        self.__cursor.executemany(sql, seq_of_parameters)
        return self

    def executescript(self, sql_script: str) -> Self:
        self.__cursor.executescript(sql_script)
        return self

    def fetchone(self) -> sqlite3.Row | tuple[Any, ...] | None:
        return self.__cursor.fetchone()

    def fetchmany(self, size: int | None = None) -> list[sqlite3.Row | tuple[Any, ...]]:
        if size is None:
            return self.__cursor.fetchmany()
        return self.__cursor.fetchmany(size)

    def fetchall(self) -> list[sqlite3.Row | tuple[Any, ...]]:
        return self.__cursor.fetchall()

    def close(self) -> None:
        self.__cursor.close()

    def __iter__(self) -> Self:
        return self

    def __next__(self) -> sqlite3.Row | tuple[Any, ...]:
        return next(self.__cursor)


class _ConnectionCapability:
    """Supported public SQLite operations without a raw native capability."""

    __slots__ = ("__native_connection",)

    def __init__(self, native_connection: Callable[[], sqlite3.Connection]) -> None:
        self.__native_connection = native_connection

    def __native(self) -> sqlite3.Connection:
        return self.__native_connection()

    @property
    def in_transaction(self) -> bool:
        return self.__native().in_transaction

    @property
    def total_changes(self) -> int:
        return self.__native().total_changes

    @property
    def _graph_write_epoch(self) -> int:
        """Read-only trust identity used by graph facades in higher layers."""
        connection = self.__native()
        value = getattr(connection, "_graph_write_epoch", None)
        if type(value) is int:
            return value
        return connection.total_changes

    @property
    def row_factory(self) -> type[sqlite3.Row]:
        return sqlite3.Row

    @row_factory.setter
    def row_factory(self, value: object) -> None:
        if value is sqlite3.Row:
            return
        raise TypeError("custom connection row factories are not supported")

    def cursor(self, factory: object | None = None) -> _CursorCapability:
        if factory is not None:
            raise TypeError("custom cursor factories are not supported")
        return _CursorCapability(self, self.__native().cursor())

    def execute(self, sql: str, parameters: Any = ()) -> _CursorCapability:
        return _CursorCapability(self, self.__native().execute(sql, parameters))

    def executemany(self, sql: str, seq_of_parameters: Any) -> _CursorCapability:
        return _CursorCapability(self, self.__native().executemany(sql, seq_of_parameters))

    def executescript(self, sql_script: str) -> _CursorCapability:
        return _CursorCapability(self, self.__native().executescript(sql_script))

    def commit(self) -> None:
        self.__native().commit()

    def rollback(self) -> None:
        self.__native().rollback()

    def close(self) -> None:
        self.__native().close()

    def set_authorizer(
        self,
        authorizer_callback: Callable[[int, str | None, str | None, str | None, str | None], int]
        | None,
    ) -> None:
        self.__native().set_authorizer(authorizer_callback)

    def set_progress_handler(self, progress_handler: Callable[[], int] | None, n: int) -> None:
        self.__native().set_progress_handler(progress_handler, n)

    def set_trace_callback(self, trace_callback: Callable[[str], object] | None) -> None:
        self.__native().set_trace_callback(trace_callback)

    def __enter__(self) -> Self:
        self.__native().__enter__()
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> bool:
        return bool(self.__native().__exit__(exc_type, exc_value, traceback))


def _unique_errors(errors: Sequence[BaseException]) -> list[BaseException]:
    """Keep cleanup evidence once per exception identity, in occurrence order."""
    seen: set[int] = set()
    unique: list[BaseException] = []
    for error in errors:
        if id(error) not in seen:
            seen.add(id(error))
            unique.append(error)
    return unique


def _normalize_exception_group(error: BaseException) -> BaseException:
    """Compatibility wrapper for identity-safe standalone cleanup evidence."""
    normalized = normalize_exception_graph(error)
    if normalized is None:
        raise AssertionError("exception evidence unexpectedly normalized to empty")
    return normalized


def _native_connection_is_closed(
    connection: sqlite3.Connection,
    cleanup_errors: list[BaseException],
) -> bool:
    """Prove physical descriptor closure without trusting virtual state."""
    try:
        descriptor = cast(Any, sqlite3.Connection.in_transaction)
        _ = descriptor.__get__(connection, sqlite3.Connection)
    except sqlite3.ProgrammingError as state_error:
        if "closed" in str(state_error).casefold():
            return True
        cleanup_errors.append(state_error)
    except BaseException as state_error:
        cleanup_errors.append(state_error)
    return False


def _connection_is_closed(
    connection: SQLiteConnectionLike,
    cleanup_errors: list[BaseException],
) -> bool:
    """Prove closure natively when possible, virtually for test facades."""
    if isinstance(connection, sqlite3.Connection):
        return _native_connection_is_closed(connection, cleanup_errors)
    try:
        _ = connection.in_transaction
    except sqlite3.ProgrammingError as state_error:
        if "closed" in str(state_error).casefold():
            return True
        cleanup_errors.append(state_error)
    except BaseException as state_error:
        cleanup_errors.append(state_error)
    return False


def _native_connection_in_transaction(
    connection: sqlite3.Connection,
    cleanup_errors: list[BaseException],
) -> bool:
    """Read transaction state from the native descriptor, failing closed."""
    try:
        descriptor = cast(Any, sqlite3.Connection.in_transaction)
        return bool(descriptor.__get__(connection, sqlite3.Connection))
    except sqlite3.ProgrammingError as state_error:
        if "closed" in str(state_error).casefold():
            return False
        cleanup_errors.append(state_error)
    except BaseException as state_error:
        cleanup_errors.append(state_error)
    return True


def _rollback_native_connection(
    connection: sqlite3.Connection,
    cleanup_errors: list[BaseException],
) -> None:
    """Release a native transaction even when virtual rollback misbehaves."""
    if not _native_connection_in_transaction(connection, cleanup_errors):
        return
    try:
        connection.rollback()
    except BaseException as rollback_error:
        cleanup_errors.append(rollback_error)
    if not _native_connection_in_transaction(connection, cleanup_errors):
        return
    try:
        sqlite3.Connection.rollback(connection)
    except BaseException as rollback_error:
        cleanup_errors.append(rollback_error)


def _conclusively_close_native_connection(
    connection: SQLiteConnectionLike,
    cleanup_errors: list[BaseException],
    *,
    virtual_close_callbacks: Sequence[Callable[[], None]],
) -> bool:
    """Attempt virtual cleanup, force native close, and prove the result."""
    for close_connection in virtual_close_callbacks:
        try:
            close_connection()
        except BaseException as close_error:
            cleanup_errors.append(close_error)
        if _connection_is_closed(connection, cleanup_errors):
            break

    # Always cross the native descriptor boundary.  Instrumented subclasses
    # may return or raise before their override reaches sqlite3.Connection.
    try:
        if isinstance(connection, sqlite3.Connection):
            sqlite3.Connection.close(connection)
        else:
            connection.close()
    except BaseException as close_error:
        cleanup_errors.append(close_error)
    return _connection_is_closed(connection, cleanup_errors)


class IndexStore:
    """One SQLite index with deterministic schema and transaction helpers."""

    def __init__(self, path: str | Path, *, prefer_rtree: bool = True) -> None:
        self.path = Path(path).expanduser().resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._connection: sqlite3.Connection = _open_index_connection(self.path)
        self._closed = False
        try:
            self._connection_capability = _ConnectionCapability(
                self._native_connection_for_capability
            )
            self.schema_created = False
            self.schema_rebuilt = False
            self._initialize_schema(prefer_rtree=prefer_rtree)
            self.edge_store = EdgeStore(self._connection)
        except BaseException as primary_error:
            cleanup_errors: list[BaseException] = []
            _rollback_native_connection(self._connection, cleanup_errors)
            self._closed = _conclusively_close_native_connection(
                self._connection,
                cleanup_errors,
                virtual_close_callbacks=(self._connection.close,),
            )
            cleanup_errors = _unique_errors(cleanup_errors)
            if not cleanup_errors:
                raise
            cleanup_failure: BaseException = cleanup_errors[0]
            if len(cleanup_errors) > 1:
                cleanup_failure = BaseExceptionGroup(
                    "IndexStore initialization cleanup failures",
                    cleanup_errors,
                )
            cleanup_failure = _normalize_exception_group(cleanup_failure)
            prepared_failure = prepare_chained_failure_with_primary_evidence(
                cleanup_failure,
                primary_error,
                message="IndexStore initialization causal evidence and cleanup failure",
            )
            primary_error.add_note(
                "IndexStore initialization cleanup also failed; the original error remains primary."
            )
            if prepared_failure is not None:
                raise primary_error from prepared_failure
            raise primary_error

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        del exc_type, traceback
        close_error: BaseException | None = None
        try:
            self.close()
        except BaseException as caught_close_error:
            close_error = caught_close_error
        if close_error is None:
            return
        if exc_value is None:
            raise close_error

        exc_value.add_note("IndexStore close also failed; the context body error remains primary.")
        close_error = prepare_chained_failure_with_primary_evidence(
            close_error,
            exc_value,
            message="IndexStore body causal evidence and close failure",
        )
        if close_error is not None:
            raise exc_value from close_error
        raise exc_value

    @property
    def connection(self) -> SQLiteConnectionLike:
        """Expose supported SQL operations without the native SQLite handle."""
        if self._closed:
            raise RuntimeError("index store is closed")
        return self._connection_capability

    def _native_connection_for_capability(self) -> sqlite3.Connection:
        # A capability retained before close preserves SQLite's native closed-
        # handle ProgrammingError.  Acquiring a capability after close remains
        # forbidden by the public property above.
        return self._connection

    @property
    def dependency_graph(self) -> DependencyGraph:
        """Return the validated block-level dependency query facade."""
        if self._closed:
            raise RuntimeError("index store is closed")
        from excel_lsp.core.graph.queries import DependencyGraph

        return DependencyGraph(self._connection, self.edge_store)

    @property
    def generation(self) -> int:
        """Return the current monotonically increasing index generation."""
        raw = self.get_meta("generation", "0")
        try:
            return int(raw or "0")
        except ValueError as exc:
            raise RuntimeError("index generation is not an integer") from exc

    def close(self) -> None:
        """Rollback abandoned work and conclusively close the connection."""
        if self._closed:
            return
        cleanup_errors: list[BaseException] = []
        transaction_state_known = True
        try:
            transaction_active = self._connection.in_transaction
        except sqlite3.ProgrammingError as state_error:
            if "closed" in str(state_error).casefold():
                transaction_active = False
            else:
                cleanup_errors.append(state_error)
                transaction_state_known = False
                transaction_active = True
        except BaseException as state_error:
            cleanup_errors.append(state_error)
            transaction_state_known = False
            transaction_active = True

        try:
            if transaction_active:
                try:
                    self._rollback_after_failed_transaction()
                except BaseException as rollback_error:
                    cleanup_errors.append(rollback_error)
                    try:
                        transaction_active = self._connection.in_transaction
                    except BaseException as state_error:
                        cleanup_errors.append(state_error)
                        transaction_state_known = False
                    if transaction_state_known and transaction_active:
                        try:
                            self._connection.execute("ROLLBACK")
                        except BaseException as fallback_error:
                            cleanup_errors.append(fallback_error)
                try:
                    self.edge_store.transaction_rolled_back()
                except BaseException as hook_error:
                    cleanup_errors.append(hook_error)
            try:
                self.edge_store.finalize_transaction_rollback()
            except BaseException as finalizer_error:
                cleanup_errors.append(finalizer_error)
        finally:
            self._closed = self._conclusively_close_connection(cleanup_errors)

        cleanup_errors = _unique_errors(cleanup_errors)
        if cleanup_errors:
            if len(cleanup_errors) == 1:
                raise _normalize_exception_group(cleanup_errors[0])
            cleanup_failure = BaseExceptionGroup("IndexStore close failures", cleanup_errors)
            raise _normalize_exception_group(cleanup_failure)

    def _close_connection(self) -> None:
        self._connection.close()

    def _conclusively_close_connection(self, cleanup_errors: list[BaseException]) -> bool:
        """Close the SQLite handle, including instrumented native subclasses."""
        return _conclusively_close_native_connection(
            self._connection,
            cleanup_errors,
            virtual_close_callbacks=(
                self._close_connection,
                self._connection.close,
                self._connection.close,
            ),
        )

    def _connection_is_closed(self, cleanup_errors: list[BaseException]) -> bool:
        """Prove physical descriptor closure without trusting close's return."""
        return _connection_is_closed(self._connection, cleanup_errors)

    @contextmanager
    def transaction(self) -> Generator[None, None, None]:
        """Run an immediate atomic mutation, nesting only in a store-owned transaction."""
        if self._connection.in_transaction:
            self.edge_store.require_store_owned_transaction()
            yield
            return
        try:
            # BEGIN may take effect before an instrumented native connection
            # raises.  Keep acquisition inside managed cleanup ownership.
            self._connection.execute("BEGIN IMMEDIATE")
            self.edge_store.transaction_started()
            yield
        except BaseException as body_error:
            self._cleanup_failed_transaction(body_error)
            raise
        else:
            try:
                self._connection.commit()
            except BaseException as commit_error:
                self._cleanup_failed_transaction(commit_error)
                raise
            self._finalize_successful_commit()

    def _rollback_after_failed_transaction(self) -> None:
        """Rollback a managed transaction after its body or commit failed."""
        self._connection.rollback()

    def _cleanup_failed_transaction(self, primary_error: BaseException) -> None:
        """Restore database and graph trust state without masking ``primary_error``."""
        cleanup_errors: list[BaseException] = []
        try:
            self._rollback_after_failed_transaction()
        except BaseException as cleanup_error:
            cleanup_errors.append(cleanup_error)
            try:
                still_active = self._connection.in_transaction
            except BaseException as state_error:
                cleanup_errors.append(state_error)
                still_active = True
            if still_active:
                try:
                    self._connection.execute("ROLLBACK")
                except BaseException as fallback_error:
                    cleanup_errors.append(fallback_error)
        transaction_state_known = True
        try:
            still_active = self._connection.in_transaction
        except BaseException as state_error:
            cleanup_errors.append(state_error)
            transaction_state_known = False
            still_active = True
        try:
            self.edge_store.transaction_rolled_back()
        except BaseException as cleanup_error:
            cleanup_errors.append(cleanup_error)
        try:
            self.edge_store.finalize_transaction_rollback()
        except BaseException as cleanup_error:
            cleanup_errors.append(cleanup_error)
        if not transaction_state_known or still_active:
            cleanup_errors.append(
                RuntimeError("managed transaction could not prove SQLite rollback")
            )
            self._closed = self._conclusively_close_connection(cleanup_errors)
        cleanup_errors = _unique_errors(cleanup_errors)
        if not cleanup_errors:
            return

        cleanup_failure: BaseException = cleanup_errors[0]
        if len(cleanup_errors) > 1:
            cleanup_failure = BaseExceptionGroup(
                "managed transaction cleanup failures",
                cleanup_errors,
            )
        prepared_failure = prepare_chained_failure_with_primary_evidence(
            cleanup_failure,
            primary_error,
            message="Managed transaction body causal evidence and cleanup failure",
        )
        primary_error.add_note(
            "Managed transaction cleanup also failed; the original error remains primary."
        )
        if prepared_failure is not None:
            raise primary_error from prepared_failure
        raise primary_error

    def _finalize_successful_commit(self) -> None:
        """Publish graph bookkeeping after SQLite has durably committed."""
        hook_error: BaseException | None = None
        try:
            self.edge_store.transaction_committed()
        except BaseException as error:
            hook_error = error

        finalizer_error: BaseException | None = None
        try:
            self.edge_store.finalize_transaction_commit()
        except BaseException as error:
            finalizer_error = error

        if hook_error is not None:
            hook_error = _normalize_exception_group(hook_error)
            if finalizer_error is not None:
                hook_error.add_note(
                    "Commit bookkeeping finalization also failed; "
                    "the commit-hook error remains primary."
                )
                finalizer_error = prepare_chained_failure_with_primary_evidence(
                    finalizer_error,
                    hook_error,
                    message="Commit-hook causal evidence and finalization failure",
                )
                if finalizer_error is not None:
                    raise hook_error from finalizer_error
                raise hook_error
            raise _normalize_exception_group(hook_error)
        if finalizer_error is not None:
            raise _normalize_exception_group(finalizer_error)

    def get_meta(self, key: str, default: str | None = None) -> str | None:
        """Read one metadata value."""
        row = self._connection.execute("SELECT value FROM meta WHERE key = ?", (key,)).fetchone()
        return default if row is None else str(row[0])

    def meta_items(self) -> dict[str, str]:
        """Return all metadata key/value pairs."""
        return {
            str(row[0]): str(row[1])
            for row in self._connection.execute(
                "SELECT key, value FROM meta ORDER BY key"
            ).fetchall()
        }

    def set_meta(self, key: str, value: object) -> None:
        """Upsert one metadata value as text."""
        self._connection.execute(
            "INSERT INTO meta(key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, str(value)),
        )

    def set_meta_many(self, values: Mapping[str, object]) -> None:
        """Upsert a group of metadata values."""
        self._connection.executemany(
            "INSERT INTO meta(key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            ((key, str(value)) for key, value in values.items()),
        )

    def bump_generation(self) -> int:
        """Increment and return the generation inside the caller's transaction."""
        generation = self.generation + 1
        self.set_meta("generation", generation)
        return generation

    def get_part_hash(self, part_name: str) -> str | None:
        """Read one normalized OOXML part hash."""
        row = self._connection.execute(
            "SELECT part_hash FROM package_parts WHERE part_name = ?",
            (_normalize_part_name(part_name),),
        ).fetchone()
        return None if row is None else str(row[0])

    def get_part_hashes(self) -> dict[str, str]:
        """Return all selected OOXML part hashes."""
        return {
            str(row[0]): str(row[1])
            for row in self._connection.execute(
                "SELECT part_name, part_hash FROM package_parts ORDER BY part_name"
            ).fetchall()
        }

    def replace_part_hashes(
        self,
        parts: Mapping[str, str],
        *,
        kinds: Mapping[str, str] | None = None,
    ) -> None:
        """Replace the selected-package hash snapshot."""
        normalized_kinds = {
            _normalize_part_name(name): kind for name, kind in (kinds or {}).items()
        }
        rows: list[tuple[str, str, str]] = []
        for raw_name, part_hash in parts.items():
            name = _normalize_part_name(raw_name)
            rows.append(
                (
                    name,
                    part_hash,
                    normalized_kinds.get(name, _part_kind(name)),
                )
            )
        self._connection.execute("DELETE FROM package_parts")
        self._connection.executemany(
            "INSERT INTO package_parts(part_name, part_hash, kind) VALUES (?, ?, ?)",
            sorted(rows),
        )

    def replace_sheet_catalog(self, sheets: Sequence[SheetDescriptor]) -> None:
        """Reset workbook-derived rows and insert sheets in workbook order."""
        ordered = sorted(sheets, key=lambda sheet: sheet.order)
        expected_orders = list(range(len(ordered)))
        if [sheet.order for sheet in ordered] != expected_orders:
            raise ValueError("sheet descriptor orders must be contiguous and zero-based")
        normalized_names = [sheet.name.casefold() for sheet in ordered]
        if len(set(normalized_names)) != len(normalized_names):
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Workbook contains duplicate sheet names.",
            )

        self.edge_store.clear()
        for table in CONTENT_TABLES_DELETE_ORDER:
            self._connection.execute(f"DELETE FROM {table}")
        self._connection.execute("DELETE FROM package_parts")

        try:
            self._connection.executemany(
                """
                INSERT INTO sheets(
                    id, name, xml_part, part_hash, kind, visibility, max_row, max_col
                ) VALUES (?, ?, ?, '', ?, ?, 0, 0)
                """,
                (
                    (
                        descriptor.order + 1,
                        descriptor.name,
                        _normalize_part_name(descriptor.xml_part),
                        descriptor.kind,
                        descriptor.visibility,
                    )
                    for descriptor in ordered
                ),
            )
        except sqlite3.IntegrityError as error:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Workbook sheet catalog violates OOXML uniqueness constraints.",
            ) from error

    def replace_defined_names(self, metadata: WorkbookMetadata) -> None:
        """Replace defined names, mapping zero-based workbook scope to DB ids."""
        self._connection.execute("DELETE FROM name_areas")
        self._connection.execute("DELETE FROM defined_names")

        db_ids_by_order: dict[int, int] = {}
        db_ids_by_name: dict[str, int] = {}
        for descriptor in metadata.sheets:
            row = self._connection.execute(
                "SELECT id FROM sheets WHERE name = ?", (descriptor.name,)
            ).fetchone()
            if row is None:
                raise ValueError(f"sheet is missing from index catalog: {descriptor.name}")
            db_id = int(row[0])
            db_ids_by_order[descriptor.order] = db_id
            db_ids_by_name[descriptor.name] = db_id

        for defined_name in metadata.defined_names:
            scope_id: int | None = None
            if defined_name.scope_sheet_order is not None:
                try:
                    scope_id = db_ids_by_order[defined_name.scope_sheet_order]
                except KeyError as exc:
                    raise ValueError(
                        "defined-name scope does not identify a workbook sheet"
                    ) from exc
            cursor = self._connection.execute(
                """
                INSERT INTO defined_names(
                    name, scope_sheet_id, refers_to, kind, is_builtin
                ) VALUES (?, ?, ?, ?, ?)
                """,
                (
                    defined_name.name,
                    scope_id,
                    defined_name.refers_to,
                    defined_name.kind,
                    int(defined_name.is_builtin),
                ),
            )
            if cursor.lastrowid is None:
                raise RuntimeError("SQLite did not return a defined-name id")
            name_id = int(cursor.lastrowid)
            area_rows: list[tuple[int, int, int, int, int, int]] = []
            for area in defined_name.areas:
                try:
                    sheet_id = db_ids_by_name[area.sheet_name]
                except KeyError as exc:
                    raise ValueError(
                        f"defined-name area uses unknown sheet: {area.sheet_name}"
                    ) from exc
                area_rows.append(
                    (
                        name_id,
                        sheet_id,
                        area.rect.row_min,
                        area.rect.row_max,
                        area.rect.col_min,
                        area.rect.col_max,
                    )
                )
            self._connection.executemany(
                """
                INSERT INTO name_areas(
                    name_id, sheet_id, row_min, row_max, col_min, col_max
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                area_rows,
            )

    def prepare_list_object_refresh(self, sheets: Sequence[SheetDescriptor]) -> None:
        """Release table aliases for one atomic multi-sheet replacement batch.

        Table names and display names are unique across a workbook. Removing
        every selected sheet's old catalog before inserting any replacement
        lets a valid table move between sheets without colliding with its stale
        owner. The caller must keep the complete replacement in one transaction
        so an intermediate catalog cannot escape if parsing or validation fails.
        """
        if not self._connection.in_transaction:
            raise RuntimeError("ListObject refresh preparation requires an active transaction")
        ordered = tuple(sorted(sheets, key=lambda sheet: sheet.order))
        if len({sheet.order for sheet in ordered}) != len(ordered):
            raise ValueError("ListObject refresh sheet selection contains duplicates")
        sheet_ids = tuple(self._sheet_id(descriptor) for descriptor in ordered)
        for sheet_id in sheet_ids:
            self._connection.execute(
                """
                DELETE FROM list_object_columns
                WHERE list_object_id IN (
                    SELECT id FROM list_objects WHERE sheet_id = ?
                )
                """,
                (sheet_id,),
            )
            self._connection.execute(
                "DELETE FROM list_objects WHERE sheet_id = ?",
                (sheet_id,),
            )

    def replace_sheet(
        self,
        descriptor: SheetDescriptor,
        parse_sheet: SheetParser,
        *,
        batch_size: int = 1_000,
        styles: StyleCatalog = DEFAULT_STYLE_CATALOG,
        region_options: RegionOptions | None = None,
    ) -> SheetParseSummary:
        """Stream one parser callback into an atomic per-sheet replacement.

        When called outside a wider transaction this method owns the mutation
        and bumps the generation. Lifecycle code wraps all changed sheets in one
        transaction and performs one aggregate generation bump.
        """
        if batch_size < 1:
            raise ValueError("batch_size must be positive")
        owns_transaction = not self._connection.in_transaction
        with self.transaction():
            sheet_id = self._upsert_sheet_descriptor(descriptor)
            self._clear_sheet_rows(sheet_id)
            pending: list[tuple[object, ...]] = []

            def flush() -> None:
                if not pending:
                    return
                try:
                    self._connection.executemany(_CELL_INSERT_SQL, pending)
                except sqlite3.IntegrityError as error:
                    raise ExcelLSPError(
                        ErrorCode.CORRUPT,
                        "Worksheet contains duplicate cell coordinates.",
                        details={"sheet": descriptor.name},
                    ) from error
                pending.clear()

            def on_cell(cell: CellRecord) -> None:
                value = _sqlite_scalar(normalize_value(cell.value), ref=cell.ref)
                pending.append(
                    (
                        sheet_id,
                        cell.row,
                        cell.col,
                        cell.ref,
                        value,
                        cell.value_type,
                        cell.formula,
                        cell.style_idx,
                        cell.formula_kind,
                        cell.shared_index,
                        cell.array_ref,
                        _data_table_json(cell.data_table),
                    )
                )
                if len(pending) >= batch_size:
                    flush()

            summary = parse_sheet(on_cell)
            flush()
            if summary.descriptor.name != descriptor.name:
                raise ValueError("sheet parser returned a summary for a different sheet")
            region_analysis = analyze_sheet_regions(
                summary,
                styles,
                lambda: self._iter_region_cells(sheet_id),
                region_options,
            )
            self._insert_region_analysis(sheet_id, region_analysis)
            self._insert_list_objects(sheet_id, summary)
            self._connection.executemany(
                """
                INSERT INTO validations(
                    sheet_id, row_min, row_max, col_min, col_max, vtype,
                    operator, formula1, formula2, allow_blank
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    (
                        sheet_id,
                        validation.rect.row_min,
                        validation.rect.row_max,
                        validation.rect.col_min,
                        validation.rect.col_max,
                        validation.validation_type,
                        validation.operator,
                        validation.formula1,
                        validation.formula2,
                        int(validation.allow_blank),
                    )
                    for validation in summary.validations
                ),
            )
            self._connection.execute(
                """
                UPDATE sheets
                SET name = ?, xml_part = ?, part_hash = ?, kind = ?, visibility = ?,
                    max_row = ?, max_col = ?
                WHERE id = ?
                """,
                (
                    descriptor.name,
                    _normalize_part_name(descriptor.xml_part),
                    summary.part_hash,
                    descriptor.kind,
                    descriptor.visibility,
                    summary.max_row,
                    summary.max_col,
                    sheet_id,
                ),
            )
            if owns_transaction:
                self.bump_generation()
            return summary

    def replace_formula_analysis(
        self,
        metadata: WorkbookMetadata,
        sheets: Sequence[SheetDescriptor] | None = None,
    ) -> tuple[str, ...]:
        """Replace P3 formula semantics for selected source sheets atomically.

        All worksheet and ListObject rows must already reflect ``metadata``.
        Calling this method without an outer transaction performs one generation
        bump; lifecycle refreshes nest it inside their aggregate transaction.
        """
        selected = tuple(
            sorted(metadata.sheets if sheets is None else sheets, key=lambda sheet: sheet.order)
        )
        _validate_formula_sheet_selection(metadata.sheets, selected)
        owns_transaction = not self._connection.in_transaction
        with self.transaction():
            context = self._formula_reference_context(metadata)
            analyzed: list[str] = []
            for descriptor in selected:
                sheet_id = self._sheet_id(descriptor)
                formula_cells = tuple(
                    FormulaCell(
                        row=int(row["row"]),
                        col=int(row["col"]),
                        formula=str(row["formula"]),
                    )
                    for row in self._connection.execute(
                        """
                        SELECT row, col, formula
                        FROM cells
                        WHERE sheet_id = ? AND formula IS NOT NULL
                        ORDER BY col, row
                        """,
                        (sheet_id,),
                    )
                )
                regions = tuple(
                    Rect(
                        int(row["row_min"]),
                        int(row["row_max"]),
                        int(row["col_min"]),
                        int(row["col_max"]),
                    )
                    for row in self._connection.execute(
                        """
                        SELECT row_min, row_max, col_min, col_max
                        FROM regions WHERE sheet_id = ? ORDER BY n
                        """,
                        (sheet_id,),
                    )
                )
                analysis = analyze_sheet_formulas(
                    descriptor,
                    formula_cells,
                    regions,
                    context,
                )
                self._replace_sheet_formula_analysis(sheet_id, analysis)
                analyzed.append(descriptor.name)
            self._replace_circular_diagnostics(metadata, context)
            if owns_transaction:
                self.bump_generation()
            # Keep the validated graph rebuild as the final write in the
            # transaction.  Its process-local live seal must never be rebased
            # over later writes at commit time.
            self._rebuild_graph_spatial_index()
            return tuple(analyzed)

    def rebuild_graph_spatial_index(self) -> None:
        """Rebuild both ranked graph mirrors from the relational edge catalog."""
        owns_transaction = not self._connection.in_transaction
        with self.transaction():
            if owns_transaction:
                self.bump_generation()
            self._rebuild_graph_spatial_index()

    def _rebuild_graph_spatial_index(self) -> None:
        """Validate all graph edges and assign dense public-hop ranks."""
        try:
            records = canonical_ranked_edges(self._connection)
            self.edge_store.rebuild_ranked(records)
        except GraphProjectionError as exc:
            detail = str(exc)
            message = detail[:1].upper() + detail[1:]
            if message and message[-1] not in ".!?":
                message += "."
            raise ExcelLSPError(ErrorCode.CORRUPT, message) from exc
        except (RuntimeError, TypeError, ValueError) as exc:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Dependency graph spatial mirrors could not be rebuilt.",
            ) from exc

    def _replace_circular_diagnostics(
        self,
        metadata: WorkbookMetadata,
        context: ReferenceContext,
    ) -> None:
        """Recompute workbook-wide P4 circular findings in the active transaction."""
        from excel_lsp.core.graph.circular import (
            BlockKey,
            CellDependency,
            CircularBlock,
            CircularEdge,
            detect_circular_references,
        )

        placeholders = ",".join("?" for _code in _P4_DIAGNOSTIC_CODES)
        self._connection.execute(
            f"DELETE FROM diagnostics WHERE code IN ({placeholders})",
            _P4_DIAGNOSTIC_CODES,
        )
        descriptors = {descriptor.order + 1: descriptor for descriptor in metadata.sheets}
        blocks: list[CircularBlock] = []
        block_keys: dict[tuple[str, int, int], BlockKey] = {}
        anchors: dict[BlockKey, tuple[SheetDescriptor, Rect, str]] = {}
        labels: dict[BlockKey, str] = {}
        owners_by_sheet: dict[int, list[CircularBlock]] = {}

        for row in self._connection.execute(
            """
            SELECT f.id, f.sheet_id, f.n, f.row_min, f.row_max,
                   f.col_min, f.col_max, c.formula
            FROM fblocks AS f
            LEFT JOIN cells AS c
              ON c.sheet_id = f.sheet_id
             AND c.row = f.row_min AND c.col = f.col_min
            ORDER BY f.sheet_id, f.n
            """
        ):
            sheet_id = int(row["sheet_id"])
            descriptor = descriptors.get(sheet_id)
            if descriptor is None:
                raise self._circular_corrupt("formula block uses an unknown sheet")
            formula = row["formula"]
            if formula is None:
                raise self._circular_corrupt("formula block anchor has no formula")
            rect = Rect(
                int(row["row_min"]),
                int(row["row_max"]),
                int(row["col_min"]),
                int(row["col_max"]),
            )
            key = BlockKey(sheet_id, int(row["n"]))
            block = CircularBlock(key, rect)
            blocks.append(block)
            block_keys[("fblock", sheet_id, int(row["id"]))] = key
            anchors[key] = (descriptor, rect, str(formula))
            labels[key] = formula_block_symbol_id(descriptor.name, key.ordinal)
            owners_by_sheet.setdefault(sheet_id, []).append(block)

        edge_rows = self._connection.execute(
            """
            SELECT id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                   dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            FROM edges ORDER BY id
            """
        ).fetchall()
        circular_edges: list[CircularEdge] = []
        for row in edge_rows:
            edge_id = int(row["id"])
            source_kind = str(row["src_kind"])
            source_id = int(row["src_id"])
            source_sheet_id = int(row["src_sheet_id"])
            source_key = block_keys.get((source_kind, source_sheet_id, source_id))
            if source_kind == "cell" and source_key is None:
                source_key = self._circular_cell_source(
                    source_id,
                    source_sheet_id,
                    descriptors,
                    blocks,
                    anchors,
                    labels,
                    owners_by_sheet,
                )
                block_keys[(source_kind, source_sheet_id, source_id)] = source_key
            if source_key is None:
                raise self._circular_corrupt(
                    f"dependency edge {edge_id} has an unsupported or orphaned source"
                )
            if source_key.sheet_id != source_sheet_id:
                raise self._circular_corrupt(
                    f"dependency edge {edge_id} source is on another sheet"
                )

            destination_values = (
                row["dst_sheet_id"],
                row["dst_row_min"],
                row["dst_row_max"],
                row["dst_col_min"],
                row["dst_col_max"],
            )
            if all(value is None for value in destination_values):
                destination_sheet_id = None
                destination_rect = None
            elif any(value is None for value in destination_values):
                raise self._circular_corrupt(f"dependency edge {edge_id} has a partial destination")
            else:
                destination_sheet_id = int(destination_values[0])
                if destination_sheet_id not in descriptors:
                    raise self._circular_corrupt(
                        f"dependency edge {edge_id} destination sheet does not exist"
                    )
                destination_rect = Rect(
                    int(destination_values[1]),
                    int(destination_values[2]),
                    int(destination_values[3]),
                    int(destination_values[4]),
                )
            circular_edges.append(
                CircularEdge(
                    source_key,
                    destination_sheet_id,
                    destination_rect,
                    str(row["via"]),
                )
            )

        owner_indexes = {
            sheet_id: _build_circular_owner_index(sheet_owners)
            for sheet_id, sheet_owners in owners_by_sheet.items()
        }

        reference_templates: dict[BlockKey, tuple[ExtractedReference, ...]] = {}
        contextual_blocks: set[BlockKey] = set()
        for key, (descriptor, rect, anchor_formula) in anchors.items():
            anchor = FormulaAnchor(
                descriptor.order,
                descriptor.name,
                rect.row_min,
                rect.col_min,
            )
            references = analyze_formula(
                anchor_formula,
                anchor=anchor,
                context=context,
            ).references
            reference_templates[key] = references
            if any(
                reference.via.startswith("structured:")
                or reference.via == "spill"
                or bool(reference.extrusion_geometries)
                for reference in references
            ):
                contextual_blocks.add(key)

        def resolve_exact(cell: CellNode) -> tuple[CellDependency, ...]:
            owner = self._circular_owner(cell, owner_indexes.get(cell.sheet_id))
            descriptor, rect, anchor_formula = anchors[owner.key]
            anchor = FormulaAnchor(
                descriptor.order,
                descriptor.name,
                cell.row,
                cell.col,
            )
            references = reference_templates[owner.key]
            if owner.key in contextual_blocks:
                formula = anchor_formula
                if (cell.row, cell.col) != (rect.row_min, rect.col_min):
                    formula = translate_a1_formula(
                        anchor_formula,
                        origin=CellRef(rect.row_min, rect.col_min),
                        target=CellRef(cell.row, cell.col),
                        preserve_coordinate_spills=True,
                    )
                references = analyze_formula(
                    formula,
                    anchor=anchor,
                    context=context,
                ).references
            dependencies: list[CellDependency] = []
            for reference in references:
                if reference.dst_sheet_order is None and reference.geometry is None:
                    dependencies.append(CellDependency(None, None, reference.via))
                    continue
                if reference.dst_sheet_order is None or reference.geometry is None:
                    raise self._circular_corrupt("formula analysis produced a partial reference")
                destination_id = reference.dst_sheet_order + 1
                if destination_id not in descriptors:
                    raise self._circular_corrupt("formula analysis produced an unknown sheet")
                dependencies.append(
                    CellDependency(
                        destination_id,
                        resolve_reference(reference.geometry, anchor.cell),
                        reference.via,
                    )
                )
            return tuple(dependencies)

        diagnostics = detect_circular_references(blocks, circular_edges, resolve_exact)
        self._connection.executemany(
            """
            INSERT INTO diagnostics(
                severity, code, sheet_id, row, col, ref, message, related
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                (
                    diagnostic.severity,
                    diagnostic.code,
                    diagnostic.ref.sheet_id,
                    diagnostic.ref.row,
                    diagnostic.ref.col,
                    cell_symbol_id(
                        descriptors[diagnostic.ref.sheet_id].name,
                        make_cell_ref(diagnostic.ref.row, diagnostic.ref.col),
                    ),
                    diagnostic.message,
                    json.dumps(
                        {
                            "candidate_blocks": [
                                labels[key] for key in diagnostic.candidate_blocks
                            ],
                            "path": [
                                cell_symbol_id(
                                    descriptors[cell.sheet_id].name,
                                    make_cell_ref(cell.row, cell.col),
                                )
                                for cell in diagnostic.related
                            ],
                        },
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ),
                )
                for diagnostic in diagnostics
            ),
        )

    def _circular_cell_source(
        self,
        packed: int,
        sheet_id: int,
        descriptors: Mapping[int, SheetDescriptor],
        blocks: list[CircularBlock],
        anchors: dict[BlockKey, tuple[SheetDescriptor, Rect, str]],
        labels: dict[BlockKey, str],
        owners_by_sheet: dict[int, list[CircularBlock]],
    ) -> BlockKey:
        from excel_lsp.core.graph.circular import BlockKey, CircularBlock

        descriptor = descriptors.get(sheet_id)
        row, col = divmod(packed, _PACKED_CELL_FACTOR)
        if descriptor is None or not 1 <= row <= 1_048_576 or not 1 <= col <= 16_384:
            raise self._circular_corrupt("packed formula-cell source is invalid")
        formula_row = self._connection.execute(
            """
            SELECT formula FROM cells
            WHERE sheet_id = ? AND row = ? AND col = ?
            """,
            (sheet_id, row, col),
        ).fetchone()
        if formula_row is None or formula_row["formula"] is None:
            raise self._circular_corrupt("formula-cell source has no formula")
        key = BlockKey(sheet_id, _PSEUDO_CELL_BLOCK_BASE + packed)
        rect = Rect(row, row, col, col)
        formula = str(formula_row["formula"])
        block = CircularBlock(key, rect)
        blocks.append(block)
        anchors[key] = (descriptor, rect, formula)
        labels[key] = cell_symbol_id(descriptor.name, make_cell_ref(row, col))
        owners_by_sheet.setdefault(sheet_id, []).append(block)
        return key

    @staticmethod
    def _circular_owner(
        cell: CellNode,
        owners: _CircularOwnerNode | None,
    ) -> CircularBlock:
        matches = _query_circular_owner_index(owners, cell.row, cell.col)
        if not matches:
            raise IndexStore._circular_corrupt("circular resolver cell has no formula source")
        return min(
            matches,
            key=lambda block: (
                block.rect != Rect(cell.row, cell.row, cell.col, cell.col),
                block.key.ordinal,
            ),
        )

    @staticmethod
    def _circular_corrupt(problem: str) -> ExcelLSPError:
        return ExcelLSPError(
            ErrorCode.CORRUPT,
            f"Dependency graph is corrupt: {problem}.",
        )

    def _formula_reference_context(self, metadata: WorkbookMetadata) -> ReferenceContext:
        # The schema's unique ``lookup_name`` column protects table names, but
        # formulas may legally use either ``name`` or ``displayName``.  Validate
        # the complete alias namespace here as a defensive boundary for
        # sidecars created by older builds or modified outside IndexStore.
        self._list_object_alias_owners()
        try:
            sheet_orders = {sheet.name: sheet.order for sheet in metadata.sheets}
            columns_by_table: dict[int, list[str]] = {}
            for row in self._connection.execute(
                """
                SELECT list_object_id, idx, name
                FROM list_object_columns
                ORDER BY list_object_id, idx
                """
            ):
                columns_by_table.setdefault(int(row["list_object_id"]), []).append(str(row["name"]))

            tables: list[TableBinding] = []
            for row in self._connection.execute(
                """
                SELECT t.id, s.name AS sheet_name, t.name, t.display_name,
                       t.row_min, t.row_max, t.col_min, t.col_max,
                       t.header_rows, t.totals_rows
                FROM list_objects AS t
                JOIN sheets AS s ON s.id = t.sheet_id
                ORDER BY s.id, t.row_min, t.col_min, t.name
                """
            ):
                sheet_name = str(row["sheet_name"])
                try:
                    sheet_order = sheet_orders[sheet_name]
                except KeyError as exc:
                    raise ValueError("ListObject catalog uses an unknown sheet") from exc
                start = make_cell_ref(int(row["row_min"]), int(row["col_min"]))
                end = make_cell_ref(int(row["row_max"]), int(row["col_max"]))
                table_id = int(row["id"])
                tables.append(
                    TableBinding(
                        sheet_order,
                        sheet_name,
                        TableInfo(
                            name=str(row["name"]),
                            display_name=str(row["display_name"]),
                            ref=start if start == end else f"{start}:{end}",
                            header_rows=int(row["header_rows"]),
                            totals_rows=int(row["totals_rows"]),
                            columns=tuple(columns_by_table.get(table_id, ())),
                        ),
                    )
                )
            return ReferenceContext(
                sheets=metadata.sheets,
                defined_names=metadata.defined_names,
                tables=tuple(tables),
                external_links=metadata.external_links,
            )
        except ValueError as error:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Workbook metadata cannot form a valid formula reference context.",
            ) from error

    def _replace_sheet_formula_analysis(
        self,
        sheet_id: int,
        analysis: SheetFormulaAnalysis,
    ) -> None:
        affected_edges = self._connection.execute(
            "SELECT id FROM edges WHERE src_sheet_id = ?",
            (sheet_id,),
        ).fetchall()
        for edge in affected_edges:
            self.edge_store.delete(int(edge[0]))
        self._connection.execute("DELETE FROM edges WHERE src_sheet_id = ?", (sheet_id,))
        self._connection.execute(
            """
            UPDATE columns SET formula_block_id = NULL
            WHERE region_id IN (SELECT id FROM regions WHERE sheet_id = ?)
            """,
            (sheet_id,),
        )
        self._connection.execute("DELETE FROM fblocks WHERE sheet_id = ?", (sheet_id,))
        placeholders = ",".join("?" for _code in _P3_DIAGNOSTIC_CODES)
        self._connection.execute(
            f"DELETE FROM diagnostics WHERE sheet_id = ? AND code IN ({placeholders})",
            (sheet_id, *_P3_DIAGNOSTIC_CODES),
        )

        block_ids: dict[int, int] = {}
        for block in analysis.blocks:
            cursor = self._connection.execute(
                """
                INSERT INTO fblocks(
                    sheet_id, n, r1c1, row_min, row_max, col_min, col_max,
                    volatile, opaque
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    sheet_id,
                    block.n,
                    block.r1c1,
                    block.rect.row_min,
                    block.rect.row_max,
                    block.rect.col_min,
                    block.rect.col_max,
                    int(block.volatile),
                    int(block.opaque),
                ),
            )
            if cursor.lastrowid is None:
                raise RuntimeError("SQLite did not return a formula-block id")
            block_ids[block.n] = int(cursor.lastrowid)

        for edge in analysis.edges:
            try:
                source_id = block_ids[edge.source_block_n]
            except KeyError as exc:
                raise RuntimeError("formula edge names an unknown source block") from exc
            destination_id = None
            if edge.dst_sheet_order is not None:
                destination_id = self._sheet_id_by_order(edge.dst_sheet_order)
            rect = edge.rect
            cursor = self._connection.execute(
                """
                INSERT INTO edges(
                    src_kind, src_id, src_sheet_id, dst_sheet_id,
                    dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
                ) VALUES ('fblock', ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    source_id,
                    sheet_id,
                    destination_id,
                    None if rect is None else rect.row_min,
                    None if rect is None else rect.row_max,
                    None if rect is None else rect.col_min,
                    None if rect is None else rect.col_max,
                    edge.via,
                ),
            )
            if cursor.lastrowid is None:
                raise RuntimeError("SQLite did not return an edge id")
            if destination_id is not None and rect is not None:
                self.edge_store.insert(int(cursor.lastrowid), destination_id, rect)

        self._connection.executemany(
            """
            INSERT INTO diagnostics(
                severity, code, sheet_id, row, col, ref, message, related
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                (
                    diagnostic.severity,
                    diagnostic.code,
                    sheet_id,
                    diagnostic.row,
                    diagnostic.col,
                    diagnostic.ref,
                    diagnostic.message,
                    json.dumps(
                        dict(sorted(diagnostic.related.items())),
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ),
                )
                for diagnostic in analysis.diagnostics
            ),
        )
        self._link_region_columns_to_blocks(sheet_id, analysis.blocks, block_ids)

    def _link_region_columns_to_blocks(
        self,
        sheet_id: int,
        blocks: Sequence[FormulaBlock],
        block_ids: Mapping[int, int],
    ) -> None:
        totals_by_table = {
            str(row["name"]).casefold(): int(row["totals_rows"])
            for row in self._connection.execute(
                "SELECT name, totals_rows FROM list_objects WHERE sheet_id = ?",
                (sheet_id,),
            )
        }
        rows = self._connection.execute(
            """
            SELECT c.id, c.idx, r.row_min, r.row_max, r.col_min,
                   r.header_rows, r.list_object_name
            FROM columns AS c
            JOIN regions AS r ON r.id = c.region_id
            WHERE r.sheet_id = ?
            ORDER BY r.n, c.idx
            """,
            (sheet_id,),
        )
        for row in rows:
            column = int(row["col_min"]) + int(row["idx"])
            body_min = int(row["row_min"]) + int(row["header_rows"])
            table_name = row["list_object_name"]
            totals_rows = (
                0 if table_name is None else totals_by_table.get(str(table_name).casefold(), 0)
            )
            body_max = int(row["row_max"]) - totals_rows
            if body_min > body_max:
                continue
            candidates = [
                block
                for block in blocks
                if block.rect.col_min <= column <= block.rect.col_max
                and block.rect.row_min <= body_min
                and block.rect.row_max >= body_max
            ]
            if len(candidates) == 1:
                self._connection.execute(
                    "UPDATE columns SET formula_block_id = ? WHERE id = ?",
                    (block_ids[candidates[0].n], int(row["id"])),
                )

    def _sheet_id(self, descriptor: SheetDescriptor) -> int:
        row = self._connection.execute(
            "SELECT id FROM sheets WHERE id = ? AND name = ?",
            (descriptor.order + 1, descriptor.name),
        ).fetchone()
        if row is None:
            raise ValueError(f"sheet is missing from index catalog: {descriptor.name}")
        return int(row["id"])

    def _sheet_id_by_order(self, sheet_order: int) -> int:
        row = self._connection.execute(
            "SELECT id FROM sheets WHERE id = ?",
            (sheet_order + 1,),
        ).fetchone()
        if row is None:
            raise ValueError(f"formula destination uses unknown sheet order: {sheet_order}")
        return int(row["id"])

    def _iter_region_cells(self, sheet_id: int) -> Iterator[RegionCell]:
        rows = self._connection.execute(
            """
            SELECT row, col, value, value_type, style_idx, formula
            FROM cells WHERE sheet_id = ? ORDER BY row, col
            """,
            (sheet_id,),
        )
        for row in rows:
            value_type = cast(CellValueType, str(row["value_type"]))
            yield RegionCell(
                row=int(row["row"]),
                col=int(row["col"]),
                value=_region_cell_value(row["value"], value_type=value_type),
                value_type=value_type,
                style_idx=int(row["style_idx"]),
                formula=None if row["formula"] is None else str(row["formula"]),
            )

    def _insert_region_analysis(self, sheet_id: int, analysis: RegionAnalysis) -> None:
        for region in analysis.regions:
            cursor = self._connection.execute(
                """
                INSERT INTO regions(
                    sheet_id, n, row_min, row_max, col_min, col_max,
                    header_rows, kind, list_object_name, confidence
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    sheet_id,
                    region.n,
                    region.rect.row_min,
                    region.rect.row_max,
                    region.rect.col_min,
                    region.rect.col_max,
                    region.header_rows,
                    region.kind,
                    region.list_object_name,
                    region.confidence,
                ),
            )
            if cursor.lastrowid is None:
                raise RuntimeError("SQLite did not return a region row id")
            region_id = cursor.lastrowid
            self._connection.executemany(
                """
                INSERT INTO columns(
                    region_id, idx, header, norm_header, dtype,
                    nonnull, distinct_est, formula_block_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, NULL)
                """,
                (
                    (
                        region_id,
                        column.idx,
                        column.header,
                        column.norm_header,
                        column.dtype,
                        column.nonnull,
                        column.distinct_est,
                    )
                    for column in region.columns
                ),
            )
        self._connection.executemany(
            """
            INSERT INTO diagnostics(
                severity, code, sheet_id, row, col, ref, message, related
            ) VALUES (?, ?, ?, NULL, NULL, ?, ?, ?)
            """,
            (
                (
                    warning.severity,
                    warning.code,
                    sheet_id,
                    warning.ref,
                    warning.message,
                    json.dumps(
                        dict(sorted(warning.related.items())),
                        ensure_ascii=False,
                        separators=(",", ":"),
                    ),
                )
                for warning in analysis.warnings
            ),
        )

    def _insert_list_objects(self, sheet_id: int, summary: SheetParseSummary) -> None:
        alias_owners = self._list_object_alias_owners()
        ordered = sorted(
            summary.tables,
            key=lambda table: (
                parse_rect(table.ref).row_min,
                parse_rect(table.ref).col_min,
                table.name.casefold(),
            ),
        )
        for table in ordered:
            aliases = (table.name, table.display_name)
            if any(not alias for alias in aliases):
                raise self._list_object_alias_error(table, summary)
            lookup_aliases = {alias.casefold() for alias in aliases}
            if lookup_aliases.intersection(alias_owners):
                raise self._list_object_alias_error(table, summary)
            rect = parse_rect(table.ref)
            try:
                cursor = self._connection.execute(
                    """
                    INSERT INTO list_objects(
                        sheet_id, name, lookup_name, display_name,
                        row_min, row_max, col_min, col_max,
                        header_rows, totals_rows
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        sheet_id,
                        table.name,
                        table.name.casefold(),
                        table.display_name,
                        rect.row_min,
                        rect.row_max,
                        rect.col_min,
                        rect.col_max,
                        table.header_rows,
                        table.totals_rows,
                    ),
                )
            except sqlite3.IntegrityError as error:
                raise ExcelLSPError(
                    ErrorCode.CORRUPT,
                    "Excel Table names and columns must be unique workbook-wide.",
                    details={"table": table.name, "sheet": summary.descriptor.name},
                ) from error
            if cursor.lastrowid is None:
                raise RuntimeError("SQLite did not return a ListObject row id")
            table_id = int(cursor.lastrowid)
            alias_owners.update({alias: table_id for alias in lookup_aliases})
            try:
                self._connection.executemany(
                    """
                    INSERT INTO list_object_columns(
                        list_object_id, idx, name, lookup_name
                    ) VALUES (?, ?, ?, ?)
                    """,
                    (
                        (table_id, index, name, name.casefold())
                        for index, name in enumerate(table.columns)
                    ),
                )
            except sqlite3.IntegrityError as error:
                raise ExcelLSPError(
                    ErrorCode.CORRUPT,
                    "Excel Table column names must be unique.",
                    details={"table": table.name, "sheet": summary.descriptor.name},
                ) from error

    def _list_object_alias_owners(self) -> dict[str, int]:
        owners: dict[str, int] = {}
        rows = self._connection.execute(
            """
            SELECT t.id, t.name, t.display_name, s.name AS sheet_name
            FROM list_objects AS t
            JOIN sheets AS s ON s.id = t.sheet_id
            ORDER BY t.id
            """
        )
        for row in rows:
            table_id = int(row["id"])
            name = str(row["name"])
            display_name = str(row["display_name"])
            for alias in {name, display_name}:
                lookup_alias = alias.casefold()
                owner = owners.get(lookup_alias)
                if not alias or (owner is not None and owner != table_id):
                    raise ExcelLSPError(
                        ErrorCode.CORRUPT,
                        "Excel Table names and display names must be unique workbook-wide.",
                        details={"table": name, "sheet": str(row["sheet_name"])},
                    )
                owners[lookup_alias] = table_id
        return owners

    @staticmethod
    def _list_object_alias_error(
        table: TableInfo,
        summary: SheetParseSummary,
    ) -> ExcelLSPError:
        return ExcelLSPError(
            ErrorCode.CORRUPT,
            "Excel Table names and display names must be unique workbook-wide.",
            details={"table": table.name, "sheet": summary.descriptor.name},
        )

    def canonical_export(self) -> dict[str, tuple[tuple[object, ...], ...]]:
        """Return content rows in natural order with local row ids projected out.

        Lifecycle-only ``generation`` and ``indexed_at`` metadata are omitted so
        a full build and an incremental build of the same workbook are directly
        comparable for invariant I2. Raw external-link targets are also omitted:
        the index retains them for reference resolution, but canonical debug and
        test projections must not disclose URL credentials, paths, or tokens.
        """
        self.edge_store.require_no_temp_shadows()
        queries = {
            "meta": """
                SELECT key, value FROM main.meta
                WHERE key NOT IN ('generation', 'indexed_at', 'external_links')
                ORDER BY key
            """,
            "sheets": """
                SELECT name, xml_part, part_hash, kind, visibility, max_row, max_col
                FROM main.sheets ORDER BY name
            """,
            "regions": """
                SELECT s.name, r.n, r.row_min, r.row_max, r.col_min, r.col_max,
                       r.header_rows, r.kind, r.list_object_name, r.confidence
                FROM main.regions AS r JOIN main.sheets AS s ON s.id = r.sheet_id
                ORDER BY s.name, r.n
            """,
            "columns": """
                SELECT s.name, r.n, c.idx, c.header, c.norm_header, c.dtype,
                       c.nonnull, c.distinct_est, fb.n
                FROM main.columns AS c
                JOIN main.regions AS r ON r.id = c.region_id
                JOIN main.sheets AS s ON s.id = r.sheet_id
                LEFT JOIN main.fblocks AS fb ON fb.id = c.formula_block_id
                ORDER BY s.name, r.n, c.idx
            """,
            "list_objects": """
                SELECT s.name, t.name, t.display_name,
                       t.row_min, t.row_max, t.col_min, t.col_max,
                       t.header_rows, t.totals_rows
                FROM main.list_objects AS t
                JOIN main.sheets AS s ON s.id = t.sheet_id
                ORDER BY s.name, t.row_min, t.col_min, t.name
            """,
            "list_object_columns": """
                SELECT s.name, t.name, c.idx, c.name
                FROM main.list_object_columns AS c
                JOIN main.list_objects AS t ON t.id = c.list_object_id
                JOIN main.sheets AS s ON s.id = t.sheet_id
                ORDER BY s.name, t.name, c.idx
            """,
            "fblocks": """
                SELECT s.name, f.n, f.r1c1, f.row_min, f.row_max, f.col_min,
                       f.col_max, f.volatile, f.opaque
                FROM main.fblocks AS f JOIN main.sheets AS s ON s.id = f.sheet_id
                ORDER BY s.name, f.n
            """,
            "defined_names": """
                SELECT d.name, scope.name, d.refers_to, d.kind, d.is_builtin
                FROM main.defined_names AS d
                LEFT JOIN main.sheets AS scope ON scope.id = d.scope_sheet_id
                ORDER BY COALESCE(scope.name, ''), d.name, d.refers_to
            """,
            "name_areas": """
                SELECT d.name, scope.name, d.refers_to, area_sheet.name,
                       a.row_min, a.row_max, a.col_min, a.col_max
                FROM main.name_areas AS a
                JOIN main.defined_names AS d ON d.id = a.name_id
                LEFT JOIN main.sheets AS scope ON scope.id = d.scope_sheet_id
                JOIN main.sheets AS area_sheet ON area_sheet.id = a.sheet_id
                ORDER BY COALESCE(scope.name, ''), d.name, d.refers_to,
                         area_sheet.name, a.row_min, a.col_min, a.row_max, a.col_max
            """,
            "validations": """
                SELECT s.name, v.row_min, v.row_max, v.col_min, v.col_max,
                       v.vtype, v.operator, v.formula1, v.formula2, v.allow_blank
                FROM main.validations AS v JOIN main.sheets AS s ON s.id = v.sheet_id
                ORDER BY s.name, v.row_min, v.col_min, v.row_max, v.col_max,
                         v.vtype, v.operator, v.formula1, v.formula2, v.allow_blank
            """,
            "edges": """
                SELECT e.src_kind, src_sheet.name,
                       CASE WHEN e.src_kind = 'fblock' THEN src_block.n ELSE e.src_id END,
                       dst_sheet.name, e.dst_row_min, e.dst_row_max,
                       e.dst_col_min, e.dst_col_max, e.via
                FROM main.edges AS e
                JOIN main.sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
                LEFT JOIN main.fblocks AS src_block
                       ON e.src_kind = 'fblock' AND src_block.id = e.src_id
                LEFT JOIN main.sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
                ORDER BY src_sheet.name, e.src_kind, 3, COALESCE(dst_sheet.name, ''),
                         e.dst_row_min, e.dst_col_min, e.dst_row_max, e.dst_col_max, e.via
            """,
            "edge_ranks": """
                SELECT e.src_kind, src_sheet.name,
                       CASE WHEN e.src_kind = 'fblock' THEN src_block.n ELSE e.src_id END,
                       dst_sheet.name, e.dst_row_min, e.dst_row_max,
                       e.dst_col_min, e.dst_col_max, e.via,
                       e.dependent_rank, e.precedent_rank
                FROM main.edges AS e
                JOIN main.sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
                LEFT JOIN main.fblocks AS src_block
                       ON e.src_kind = 'fblock' AND src_block.id = e.src_id
                LEFT JOIN main.sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
                ORDER BY src_sheet.name, e.src_kind, 3, COALESCE(dst_sheet.name, ''),
                         e.dst_row_min, e.dst_col_min, e.dst_row_max, e.dst_col_max,
                         e.via, e.dependent_rank, e.precedent_rank
            """,
            "graph_spatial_state": """
                SELECT dirty, dependent_rank_max, precedent_rank_max
                FROM main.graph_spatial_state WHERE singleton = 1
            """,
            "graph_rank_keys": """
                SELECT direction, rank, key_text
                FROM main.graph_rank_keys ORDER BY direction, rank
            """,
            "diagnostics": """
                SELECT d.severity, d.code, s.name, d.row, d.col, d.ref,
                       d.message, d.related
                FROM main.diagnostics AS d JOIN main.sheets AS s ON s.id = d.sheet_id
                ORDER BY s.name, COALESCE(d.row, -1), COALESCE(d.col, -1),
                         d.code, d.ref, d.message
            """,
            "staleness": """
                SELECT s.name, st.row_min, st.row_max, st.col_min, st.col_max, st.since
                FROM main.staleness AS st JOIN main.sheets AS s ON s.id = st.sheet_id
                ORDER BY s.name, st.row_min, st.col_min, st.row_max, st.col_max, st.since
            """,
            "cells": """
                SELECT s.name, c.row, c.col, c.ref, c.value, c.value_type,
                       c.formula, c.style_idx, c.formula_kind,
                       c.shared_index, c.array_ref, c.data_table
                FROM main.cells AS c JOIN main.sheets AS s ON s.id = c.sheet_id
                ORDER BY s.name, c.row, c.col
            """,
            "package_parts": """
                SELECT part_name, part_hash, kind
                FROM main.package_parts ORDER BY part_name
            """,
        }
        if self.edge_store.backend == "rtree":
            queries["edge_rtree"] = """
                SELECT e.src_kind, src_sheet.name,
                       CASE WHEN e.src_kind = 'fblock' THEN src_block.n ELSE e.src_id END,
                       dst_sheet.name, e.dst_row_min, e.dst_row_max,
                       e.dst_col_min, e.dst_col_max, e.via,
                       spatial_min_sheet.name, spatial_max_sheet.name,
                       r.row_min, r.row_max,
                       r.col_min, r.col_max, r.rank_min, r.rank_max
                FROM main.edge_rtree AS r
                JOIN main.edges AS e ON e.id = r.edge_id
                JOIN main.sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
                LEFT JOIN main.fblocks AS src_block
                       ON e.src_kind = 'fblock' AND src_block.id = e.src_id
                LEFT JOIN main.sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
                JOIN main.sheets AS spatial_min_sheet
                     ON spatial_min_sheet.id = CAST(r.sheet_min AS INTEGER)
                JOIN main.sheets AS spatial_max_sheet
                     ON spatial_max_sheet.id = CAST(r.sheet_max AS INTEGER)
                ORDER BY src_sheet.name, e.src_kind, 3, COALESCE(dst_sheet.name, ''),
                         e.dst_row_min, e.dst_col_min, e.dst_row_max, e.dst_col_max,
                         e.via, spatial_min_sheet.name, spatial_max_sheet.name,
                         r.row_min, r.col_min, r.row_max, r.col_max,
                         r.rank_min, r.rank_max
            """
            queries["edge_source_rtree"] = """
                SELECT e.src_kind, src_sheet.name,
                       CASE WHEN e.src_kind = 'fblock' THEN src_block.n ELSE e.src_id END,
                       dst_sheet.name, e.dst_row_min, e.dst_row_max,
                       e.dst_col_min, e.dst_col_max, e.via,
                       spatial_min_sheet.name, spatial_max_sheet.name,
                       r.row_min, r.row_max, r.col_min, r.col_max,
                       r.rank_min, r.rank_max
                FROM main.edge_source_rtree AS r
                JOIN main.edges AS e ON e.id = r.edge_id
                JOIN main.sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
                LEFT JOIN main.fblocks AS src_block
                       ON e.src_kind = 'fblock' AND src_block.id = e.src_id
                LEFT JOIN main.sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
                JOIN main.sheets AS spatial_min_sheet
                     ON spatial_min_sheet.id = CAST(r.sheet_min AS INTEGER)
                JOIN main.sheets AS spatial_max_sheet
                     ON spatial_max_sheet.id = CAST(r.sheet_max AS INTEGER)
                ORDER BY src_sheet.name, e.src_kind, 3, COALESCE(dst_sheet.name, ''),
                         e.dst_row_min, e.dst_col_min, e.dst_row_max, e.dst_col_max,
                         e.via, spatial_min_sheet.name, spatial_max_sheet.name,
                         r.row_min, r.col_min, r.row_max, r.col_max,
                         r.rank_min, r.rank_max
            """
        else:
            queries["edge_intervals"] = """
                SELECT e.src_kind, src_sheet.name,
                       CASE WHEN e.src_kind = 'fblock' THEN src_block.n ELSE e.src_id END,
                       dst_sheet.name, e.dst_row_min, e.dst_row_max,
                       e.dst_col_min, e.dst_col_max, e.via,
                       spatial_sheet.name, i.row_min, i.row_max, i.col_min, i.col_max,
                       i.rank
                FROM main.edge_intervals AS i
                JOIN main.edges AS e ON e.id = i.edge_id
                JOIN main.sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
                LEFT JOIN main.fblocks AS src_block
                       ON e.src_kind = 'fblock' AND src_block.id = e.src_id
                LEFT JOIN main.sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
                JOIN main.sheets AS spatial_sheet ON spatial_sheet.id = i.sheet_id
                ORDER BY src_sheet.name, e.src_kind, 3, COALESCE(dst_sheet.name, ''),
                         e.dst_row_min, e.dst_col_min, e.dst_row_max, e.dst_col_max,
                         e.via, spatial_sheet.name,
                         i.row_min, i.col_min, i.row_max, i.col_max, i.rank
            """
            queries["edge_source_intervals"] = """
                SELECT e.src_kind, src_sheet.name,
                       CASE WHEN e.src_kind = 'fblock' THEN src_block.n ELSE e.src_id END,
                       dst_sheet.name, e.dst_row_min, e.dst_row_max,
                       e.dst_col_min, e.dst_col_max, e.via,
                       spatial_sheet.name, i.row_min, i.row_max, i.col_min, i.col_max,
                       i.rank
                FROM main.edge_source_intervals AS i
                JOIN main.edges AS e ON e.id = i.edge_id
                JOIN main.sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
                LEFT JOIN main.fblocks AS src_block
                       ON e.src_kind = 'fblock' AND src_block.id = e.src_id
                LEFT JOIN main.sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
                JOIN main.sheets AS spatial_sheet ON spatial_sheet.id = i.sheet_id
                ORDER BY src_sheet.name, e.src_kind, 3, COALESCE(dst_sheet.name, ''),
                         e.dst_row_min, e.dst_col_min, e.dst_row_max, e.dst_col_max,
                         e.via, spatial_sheet.name,
                         i.row_min, i.col_min, i.row_max, i.col_max, i.rank
            """

        return {
            name: tuple(tuple(row) for row in self._connection.execute(sql).fetchall())
            for name, sql in queries.items()
        }

    def _initialize_schema(self, *, prefer_rtree: bool) -> None:
        self._connection.execute("PRAGMA foreign_keys = OFF")
        completed = False
        try:
            for attempt in range(len(_INITIALIZATION_RETRY_DELAYS) + 1):
                self.schema_created = False
                self.schema_rebuilt = False
                try:
                    self._connection.execute("BEGIN IMMEDIATE")
                    self._initialize_schema_locked(prefer_rtree=prefer_rtree)
                    self._connection.commit()
                    completed = True
                    return
                except sqlite3.OperationalError as error:
                    self._rollback_schema_initialization(error)
                    if not _is_busy_error(error) or attempt == len(_INITIALIZATION_RETRY_DELAYS):
                        raise
                    time.sleep(_INITIALIZATION_RETRY_DELAYS[attempt])
                except _DatabaseRecreationRequired as recovery:
                    self._rollback_schema_initialization(recovery)
                    self._recreate_database(
                        initial_generation=recovery.initial_generation,
                        prefer_rtree=prefer_rtree,
                    )
                    self.schema_created = False
                    self.schema_rebuilt = True
                    completed = True
                    return
                except BaseException as error:
                    self._rollback_schema_initialization(error)
                    raise
        finally:
            # On terminal failure the constructor owns conclusive rollback and
            # close.  A best-effort PRAGMA here must never mask that primary.
            if completed and not self._closed:
                self._connection.execute("PRAGMA foreign_keys = ON")

    def _rollback_schema_initialization(self, primary_error: BaseException) -> None:
        """Rollback one failed schema attempt without replacing its primary."""
        cleanup_errors: list[BaseException] = []
        _rollback_native_connection(self._connection, cleanup_errors)
        cleanup_errors = _unique_errors(cleanup_errors)
        if not cleanup_errors:
            return
        cleanup_failure: BaseException = cleanup_errors[0]
        if len(cleanup_errors) > 1:
            cleanup_failure = BaseExceptionGroup(
                "schema initialization rollback failures",
                cleanup_errors,
            )
        cleanup_failure = _normalize_exception_group(cleanup_failure)
        prepared_failure = prepare_chained_failure_with_primary_evidence(
            cleanup_failure,
            primary_error,
            message="Schema initialization causal evidence and rollback failure",
        )
        primary_error.add_note(
            "Schema initialization rollback also failed; the original error remains primary."
        )
        if prepared_failure is not None:
            raise primary_error from prepared_failure
        raise primary_error

    def _initialize_schema_locked(self, *, prefer_rtree: bool) -> None:
        if not self._table_exists("meta"):
            existing_tables = self._user_tables()
            if existing_tables:
                self._drop_all_objects_or_recreate(initial_generation=0)
                self.schema_rebuilt = True
            self._create_schema(initial_generation=0, prefer_rtree=prefer_rtree)
            self.schema_created = True
            return

        schema_version = self.get_meta("schema_version")
        if schema_version == SCHEMA_VERSION:
            try:
                EdgeStore.ensure_schema(self._connection, prefer_rtree=prefer_rtree)
            except EdgeSchemaError:
                previous_generation = self.generation
                self._drop_all_objects_or_recreate(initial_generation=previous_generation + 1)
                self._create_schema(
                    initial_generation=previous_generation + 1,
                    prefer_rtree=prefer_rtree,
                )
                self.schema_rebuilt = True
            return

        previous_generation = self.generation
        self._drop_all_objects_or_recreate(initial_generation=previous_generation + 1)
        self._create_schema(
            initial_generation=previous_generation + 1,
            prefer_rtree=prefer_rtree,
        )
        self.schema_rebuilt = True

    def _create_schema(self, *, initial_generation: int, prefer_rtree: bool) -> None:
        _execute_script(self._connection, BASE_SCHEMA_SQL)
        EdgeStore.ensure_schema(self._connection, prefer_rtree=prefer_rtree)
        self._connection.executemany(
            "INSERT INTO meta(key, value) VALUES (?, ?)",
            (
                ("schema_version", SCHEMA_VERSION),
                ("generation", str(initial_generation)),
            ),
        )
        self._connection.execute(f"PRAGMA user_version = {int(SCHEMA_VERSION)}")

    def _drop_all_objects_or_recreate(self, *, initial_generation: int) -> None:
        try:
            self._drop_all_objects()
        except sqlite3.DatabaseError as exc:
            raise _DatabaseRecreationRequired(initial_generation) from exc

    def _recreate_database(self, *, initial_generation: int, prefer_rtree: bool) -> None:
        """Replace storage that SQLite cannot tear down through valid DDL."""
        self._close_recreation_connection(rollback=True)
        descriptor, raw_replacement_path = tempfile.mkstemp(
            prefix=f".{self.path.name}.",
            suffix=".rebuild",
            dir=self.path.parent,
        )
        descriptor_owned = True
        replacement_path: Path | None = None
        try:
            replacement_path = Path(raw_replacement_path)
            os.close(descriptor)
            descriptor_owned = False
            self._connection = _open_index_connection(replacement_path)
            self._closed = False
            self._connection.execute("PRAGMA foreign_keys = OFF")
            self._connection.execute("BEGIN IMMEDIATE")
            self._create_schema(
                initial_generation=initial_generation,
                prefer_rtree=prefer_rtree,
            )
            self._connection.commit()
            self._connection.execute("PRAGMA wal_checkpoint(TRUNCATE)").fetchone()
            self._close_recreation_connection(rollback=False)

            _remove_sqlite_sidecars(replacement_path)
            _remove_sqlite_sidecars(self.path)
            os.replace(replacement_path, self.path)
            self._connection = _open_index_connection(self.path)
            self._closed = False
        except BaseException as primary_error:
            cleanup_errors: list[BaseException] = []
            if not self._closed:
                _rollback_native_connection(self._connection, cleanup_errors)
                self._closed = _conclusively_close_native_connection(
                    self._connection,
                    cleanup_errors,
                    virtual_close_callbacks=(self._connection.close,),
                )
            cleanup_errors = _unique_errors(cleanup_errors)
            if not cleanup_errors:
                raise
            cleanup_failure: BaseException = cleanup_errors[0]
            if len(cleanup_errors) > 1:
                cleanup_failure = BaseExceptionGroup(
                    "database recreation cleanup failures",
                    cleanup_errors,
                )
            cleanup_failure = _normalize_exception_group(cleanup_failure)
            prepared_failure = prepare_chained_failure_with_primary_evidence(
                cleanup_failure,
                primary_error,
                message="Database recreation causal evidence and cleanup failure",
            )
            primary_error.add_note(
                "Database recreation cleanup also failed; the original error remains primary."
            )
            if prepared_failure is not None:
                raise primary_error from prepared_failure
            raise primary_error
        finally:
            if descriptor_owned:
                with suppress(Exception):
                    os.close(descriptor)
            with suppress(Exception):
                _remove_sqlite_artifacts_best_effort(replacement_path or Path(raw_replacement_path))

    def _close_recreation_connection(self, *, rollback: bool) -> None:
        """Conclude ownership of one recreation handle before replacing it."""
        cleanup_errors: list[BaseException] = []
        if rollback:
            _rollback_native_connection(self._connection, cleanup_errors)
        self._closed = _conclusively_close_native_connection(
            self._connection,
            cleanup_errors,
            virtual_close_callbacks=(self._connection.close,),
        )
        cleanup_errors = _unique_errors(cleanup_errors)
        if not cleanup_errors:
            return
        if len(cleanup_errors) == 1:
            raise _normalize_exception_group(cleanup_errors[0])
        raise _normalize_exception_group(
            BaseExceptionGroup("database recreation close failures", cleanup_errors)
        )

    def _drop_all_objects(self) -> None:
        # Remove triggers before virtual tables. A malformed current-version
        # sidecar can leave trigger bodies referencing an absent state table;
        # SQLite otherwise evaluates those bodies while dropping R*Tree rows
        # and leaks ``OperationalError`` instead of allowing a clean rebuild.
        rows = self._connection.execute(
            "SELECT name FROM sqlite_master WHERE type = 'trigger' AND name NOT LIKE 'sqlite_%'"
        ).fetchall()
        for row in rows:
            name = _quote_identifier(str(row[0]))
            self._connection.execute(f"DROP TRIGGER {name}")
        if self._table_exists(EdgeStore.RTREE_TABLE):
            self._connection.execute("DROP TABLE edge_rtree")
        if self._table_exists(EdgeStore.SOURCE_RTREE_TABLE):
            self._connection.execute("DROP TABLE edge_source_rtree")
        for object_type in ("view",):
            rows = self._connection.execute(
                "SELECT name FROM sqlite_master WHERE type = ? AND name NOT LIKE 'sqlite_%'",
                (object_type,),
            ).fetchall()
            for row in rows:
                name = _quote_identifier(str(row[0]))
                self._connection.execute(f"DROP {object_type.upper()} {name}")
        for table_name in self._user_tables():
            name = _quote_identifier(table_name)
            self._connection.execute(f"DROP TABLE {name}")

    def _upsert_sheet_descriptor(self, descriptor: SheetDescriptor) -> int:
        sheet_id = descriptor.order + 1
        self._connection.execute(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (?, ?, ?, '', ?, ?, 0, 0)
            ON CONFLICT(id) DO UPDATE SET
                name = excluded.name,
                xml_part = excluded.xml_part,
                kind = excluded.kind,
                visibility = excluded.visibility
            """,
            (
                sheet_id,
                descriptor.name,
                _normalize_part_name(descriptor.xml_part),
                descriptor.kind,
                descriptor.visibility,
            ),
        )
        return sheet_id

    def _clear_sheet_rows(self, sheet_id: int) -> None:
        affected_edges = self._connection.execute(
            "SELECT id FROM edges WHERE src_sheet_id = ?",
            (sheet_id,),
        ).fetchall()
        for edge in affected_edges:
            self.edge_store.delete(int(edge[0]))
        self._connection.execute(
            "DELETE FROM edges WHERE src_sheet_id = ?",
            (sheet_id,),
        )
        for table in (
            "diagnostics",
            "staleness",
            "columns",
            "regions",
            "fblocks",
            "list_object_columns",
            "list_objects",
            "validations",
            "cells",
        ):
            if table == "columns":
                self._connection.execute(
                    "DELETE FROM columns WHERE region_id IN "
                    "(SELECT id FROM regions WHERE sheet_id = ?)",
                    (sheet_id,),
                )
            elif table == "list_object_columns":
                self._connection.execute(
                    "DELETE FROM list_object_columns WHERE list_object_id IN "
                    "(SELECT id FROM list_objects WHERE sheet_id = ?)",
                    (sheet_id,),
                )
            else:
                self._connection.execute(f"DELETE FROM {table} WHERE sheet_id = ?", (sheet_id,))

    def _table_exists(self, table_name: str) -> bool:
        row = self._connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
            (table_name,),
        ).fetchone()
        return row is not None

    def _user_tables(self) -> list[str]:
        return [
            str(row[0])
            for row in self._connection.execute(
                "SELECT name FROM sqlite_master "
                "WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
            ).fetchall()
        ]


def _build_circular_owner_index(
    blocks: Sequence[CircularBlock],
) -> _CircularOwnerNode | None:
    if not blocks:
        return None
    bounds = Rect(
        min(block.rect.row_min for block in blocks),
        max(block.rect.row_max for block in blocks),
        min(block.rect.col_min for block in blocks),
        max(block.rect.col_max for block in blocks),
    )
    if len(blocks) <= 8:
        return _CircularOwnerNode(
            bounds,
            tuple(
                sorted(
                    blocks,
                    key=lambda block: (
                        block.rect.row_min,
                        block.rect.col_min,
                        block.rect.row_max,
                        block.rect.col_max,
                        block.key.ordinal,
                    ),
                )
            ),
        )
    row_span = bounds.row_max - bounds.row_min
    col_span = bounds.col_max - bounds.col_min
    if row_span >= col_span:
        ordered = sorted(
            blocks,
            key=lambda block: (
                block.rect.row_min + block.rect.row_max,
                block.rect.col_min + block.rect.col_max,
                block.key.ordinal,
            ),
        )
    else:
        ordered = sorted(
            blocks,
            key=lambda block: (
                block.rect.col_min + block.rect.col_max,
                block.rect.row_min + block.rect.row_max,
                block.key.ordinal,
            ),
        )
    middle = len(ordered) // 2
    return _CircularOwnerNode(
        bounds,
        left=_build_circular_owner_index(ordered[:middle]),
        right=_build_circular_owner_index(ordered[middle:]),
    )


def _query_circular_owner_index(
    node: _CircularOwnerNode | None,
    row: int,
    col: int,
) -> tuple[CircularBlock, ...]:
    if node is None or not (
        node.bounds.row_min <= row <= node.bounds.row_max
        and node.bounds.col_min <= col <= node.bounds.col_max
    ):
        return ()
    if node.entries:
        return tuple(
            block
            for block in node.entries
            if block.rect.row_min <= row <= block.rect.row_max
            and block.rect.col_min <= col <= block.rect.col_max
        )
    return (
        *_query_circular_owner_index(node.left, row, col),
        *_query_circular_owner_index(node.right, row, col),
    )


def _normalize_part_name(part_name: str) -> str:
    return part_name.replace("\\", "/").lstrip("/")


def _validate_formula_sheet_selection(
    workbook_sheets: Sequence[SheetDescriptor],
    selected: Sequence[SheetDescriptor],
) -> None:
    by_order = {sheet.order: sheet for sheet in workbook_sheets}
    if len(by_order) != len(workbook_sheets):
        raise ValueError("workbook sheet orders must be unique")
    seen: set[int] = set()
    for descriptor in selected:
        if descriptor.order in seen:
            raise ValueError("formula-analysis sheet selection contains duplicates")
        seen.add(descriptor.order)
        if by_order.get(descriptor.order) != descriptor:
            raise ValueError("formula-analysis sheet selection is not from workbook metadata")


def _part_kind(part_name: str) -> str:
    if part_name == "xl/workbook.xml":
        return "workbook"
    if part_name == "xl/_rels/workbook.xml.rels":
        return "workbook_rels"
    if part_name == "xl/sharedStrings.xml":
        return "shared_strings"
    if part_name == "xl/styles.xml":
        return "styles"
    if part_name.startswith("xl/externalLinks/"):
        return "external_link"
    if part_name.startswith("xl/tables/"):
        return "worksheet_metadata"
    if part_name.startswith("xl/worksheets/"):
        return "worksheet"
    return "package"


def _retry_on_busy(operation: Callable[[], object]) -> object:
    for attempt in range(len(_INITIALIZATION_RETRY_DELAYS) + 1):
        try:
            return operation()
        except sqlite3.OperationalError as error:
            if not _is_busy_error(error) or attempt == len(_INITIALIZATION_RETRY_DELAYS):
                raise
            time.sleep(_INITIALIZATION_RETRY_DELAYS[attempt])
    raise AssertionError("unreachable SQLite initialization retry state")


def _open_index_connection(path: Path) -> _TrackedConnection:
    connection = sqlite3.connect(
        path,
        timeout=5.0,
        isolation_level=None,
        cached_statements=0,
        factory=_TrackedConnection,
    )
    try:
        connection.install_graph_tracker()
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA busy_timeout = 5000")
        _retry_on_busy(lambda: connection.execute("PRAGMA journal_mode = WAL").fetchone())
        _retry_on_busy(lambda: connection.execute("PRAGMA synchronous = NORMAL").fetchone())
        connection.execute("PRAGMA foreign_keys = ON")
    except BaseException as primary_error:
        cleanup_errors: list[BaseException] = []
        _rollback_native_connection(connection, cleanup_errors)
        _conclusively_close_native_connection(
            connection,
            cleanup_errors,
            virtual_close_callbacks=(connection.close,),
        )
        cleanup_errors = _unique_errors(cleanup_errors)
        if not cleanup_errors:
            raise
        cleanup_failure: BaseException = cleanup_errors[0]
        if len(cleanup_errors) > 1:
            cleanup_failure = BaseExceptionGroup(
                "index connection configuration cleanup failures",
                cleanup_errors,
            )
        cleanup_failure = _normalize_exception_group(cleanup_failure)
        prepared_failure = prepare_chained_failure_with_primary_evidence(
            cleanup_failure,
            primary_error,
            message="Index connection configuration causal evidence and cleanup failure",
        )
        primary_error.add_note(
            "Index connection configuration cleanup also failed; "
            "the original error remains primary."
        )
        if prepared_failure is not None:
            raise primary_error from prepared_failure
        raise primary_error
    return connection


def _remove_sqlite_sidecars(path: Path) -> None:
    for suffix in ("-wal", "-shm"):
        Path(f"{path}{suffix}").unlink(missing_ok=True)


def _remove_sqlite_artifacts_best_effort(path: Path) -> None:
    for suffix in ("-wal", "-shm", ""):
        with suppress(OSError):
            Path(f"{path}{suffix}").unlink(missing_ok=True)


def _is_busy_error(error: sqlite3.OperationalError) -> bool:
    message = str(error).casefold()
    return "locked" in message or "busy" in message


def _execute_script(connection: sqlite3.Connection, script: str) -> None:
    """Execute DDL statements without sqlite3.executescript's implicit commit."""
    pending: list[str] = []
    for line in script.splitlines(keepends=True):
        pending.append(line)
        statement = "".join(pending)
        if not sqlite3.complete_statement(statement):
            continue
        if statement.strip():
            connection.execute(statement)
        pending.clear()
    if "".join(pending).strip():
        raise RuntimeError("schema SQL ended with an incomplete statement")


def _quote_identifier(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def _sqlite_scalar(value: JsonScalar, *, ref: str) -> JsonScalar:
    """Fit JSON numerics to SQLite while preserving Excel's double domain."""
    if isinstance(value, bool) or not isinstance(value, int):
        return value
    if _SQLITE_INT_MIN <= value <= _SQLITE_INT_MAX:
        return value
    try:
        real_value = float(value)
    except OverflowError as error:
        raise ExcelLSPError(
            ErrorCode.CORRUPT,
            "Worksheet contains a numeric value outside Excel's finite range.",
            details={"ref": ref},
        ) from error
    if not math.isfinite(real_value):
        raise ExcelLSPError(
            ErrorCode.CORRUPT,
            "Worksheet contains a numeric value outside Excel's finite range.",
            details={"ref": ref},
        )
    return real_value


def _region_cell_value(value: object, *, value_type: CellValueType) -> CellScalar:
    if value is None:
        return None
    if value_type == "bool":
        return bool(value)
    if isinstance(value, (int, float, str)):
        return value
    raise ExcelLSPError(
        ErrorCode.CORRUPT,
        "Index contains an unsupported stored cell value.",
        details={"valueType": type(value).__name__},
    )


def _data_table_json(data_table: DataTableFormulaInfo | None) -> str | None:
    if data_table is None:
        return None
    return json.dumps(
        {
            "calculateAlways": data_table.calculate_always,
            "deletedColumnInput": data_table.deleted_column_input,
            "deletedRowInput": data_table.deleted_row_input,
            "inputCell1": data_table.input_cell_1,
            "inputCell2": data_table.input_cell_2,
            "is2D": data_table.is_2d,
            "ref": data_table.ref,
            "rowOriented": data_table.row_oriented,
        },
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    )


__all__ = ["CellConsumer", "IndexStore", "SheetParser"]
