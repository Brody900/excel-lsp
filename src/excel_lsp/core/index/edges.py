"""Spatial range storage with an SQLite R*Tree and interval-table fallback."""

from __future__ import annotations

import json
import sqlite3
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from typing import Any, Literal, Protocol, Self, cast

from excel_lsp.core.exception_evidence import (
    normalize_exception_graph,
    prepare_chained_failure_with_primary_evidence,
)
from excel_lsp.core.models import Rect
from excel_lsp.core.parse.coordinates import make_cell_ref
from excel_lsp.core.symbols import cell_symbol_id, formula_block_symbol_id


class SQLiteConnectionLike(Protocol):
    """SQLite operations shared by native handles and public capabilities."""

    @property
    def in_transaction(self) -> bool: ...

    @property
    def total_changes(self) -> int: ...

    @property
    def row_factory(self) -> Any: ...

    @row_factory.setter
    def row_factory(self, value: Any) -> None: ...

    def cursor(self, factory: Any | None = None, /) -> Any: ...

    def execute(self, sql: str, parameters: Any = (), /) -> Any: ...

    def executemany(self, sql: str, seq_of_parameters: Any, /) -> Any: ...

    def executescript(self, sql_script: str, /) -> Any: ...

    def commit(self) -> None: ...

    def rollback(self) -> None: ...

    def close(self) -> None: ...

    def set_authorizer(self, authorizer_callback: Any, /) -> None: ...

    def set_progress_handler(self, progress_handler: Any, n: int, /) -> None: ...

    def set_trace_callback(self, trace_callback: Any, /) -> None: ...

    def __enter__(self) -> Self: ...

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: Any,
        /,
    ) -> bool | None: ...


EdgeBackend = Literal["rtree", "interval"]
EdgeDirection = Literal["dependents", "precedents"]
_GRAPH_RANK_MAX = (1 << 31) - 1
_PACKED_CELL_FACTOR = 1 << 16


class EdgeSchemaError(RuntimeError):
    """The ranked spatial backend is absent, partial, or mixed."""


class GraphProjectionError(ValueError):
    """A relational dependency cannot produce the canonical public graph."""


@dataclass(frozen=True, slots=True)
class RankedEdge:
    """One validated relational edge projected into both spatial mirrors."""

    edge_id: int
    source_sheet_id: int
    source_rect: Rect
    destination_sheet_id: int | None
    destination_rect: Rect | None
    dependent_rank: int
    precedent_rank: int
    dependent_key: tuple[object, ...]
    precedent_key: tuple[object, ...]


def canonical_ranked_edges(
    connection: SQLiteConnectionLike,
    *,
    require_persisted_ranks: bool = False,
) -> tuple[RankedEdge, ...]:
    """Project relational edges into the one canonical dense public-hop order."""
    try:
        cursor = connection.execute(
            """
            SELECT e.id, e.src_kind, e.src_id, e.src_sheet_id,
                   e.dst_sheet_id, e.dst_row_min, e.dst_row_max,
                   e.dst_col_min, e.dst_col_max, e.via,
                   e.dependent_rank, e.precedent_rank,
                   src_sheet.name AS src_sheet_name,
                   fb.id AS fblock_id, fb.sheet_id AS fblock_sheet_id,
                   fb.n AS fblock_n, fb.row_min AS src_row_min,
                   fb.row_max AS src_row_max, fb.col_min AS src_col_min,
                   fb.col_max AS src_col_max,
                   dst_sheet.name AS dst_sheet_name
            FROM main.edges AS e
            LEFT JOIN main.sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
            LEFT JOIN main.fblocks AS fb
                   ON e.src_kind = 'fblock' AND fb.id = e.src_id
            LEFT JOIN main.sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
            ORDER BY e.id
            """
        )
        names = tuple(item[0] for item in cursor.description or ())
        rows = tuple(dict(zip(names, tuple(row), strict=True)) for row in cursor)
    except sqlite3.DatabaseError as exc:
        raise GraphProjectionError("relational dependency projection could not be read") from exc

    projections: list[
        tuple[
            int,
            int,
            Rect,
            int | None,
            Rect | None,
            tuple[object, ...],
            tuple[object, ...],
            object,
            object,
        ]
    ] = []
    for row in rows:
        edge_id = _projection_integral(row["id"], 0, "edge id")
        source_sheet_id = _projection_integral(row["src_sheet_id"], edge_id, "source sheet id")
        if row["src_sheet_name"] is None:
            raise _projection_corrupt(edge_id, "source sheet does not exist")
        source_sheet = str(row["src_sheet_name"])
        if not source_sheet:
            raise _projection_corrupt(edge_id, "source sheet name is empty")
        source_kind = str(row["src_kind"])
        if source_kind == "fblock":
            if row["fblock_id"] is None:
                raise _projection_corrupt(edge_id, "formula-block source is orphaned")
            if (
                _projection_integral(row["fblock_sheet_id"], edge_id, "formula-block sheet id")
                != source_sheet_id
            ):
                raise _projection_corrupt(edge_id, "formula-block source is on another sheet")
            source_rect = _projection_rect(
                edge_id,
                row["src_row_min"],
                row["src_row_max"],
                row["src_col_min"],
                row["src_col_max"],
                "source",
            )
            block_n = _projection_integral(row["fblock_n"], edge_id, "formula-block ordinal")
            if block_n < 0:
                raise _projection_corrupt(edge_id, "formula-block ordinal is negative")
            source_label = formula_block_symbol_id(source_sheet, block_n)
        elif source_kind == "cell":
            packed = _projection_integral(row["src_id"], edge_id, "packed cell source")
            source_row, source_col = divmod(packed, _PACKED_CELL_FACTOR)
            if not 1 <= source_row <= 1_048_576 or not 1 <= source_col <= 16_384:
                raise _projection_corrupt(edge_id, "packed cell source is outside worksheet bounds")
            source_rect = Rect(source_row, source_row, source_col, source_col)
            source_label = cell_symbol_id(source_sheet, make_cell_ref(source_row, source_col))
        else:
            raise _projection_corrupt(edge_id, f"unsupported source kind {source_kind!r}")

        via = str(row["via"])
        source_key = _canonical_graph_rank_key(
            sheet=source_sheet,
            rect=source_rect,
            kind=source_kind,
            label=source_label,
            via=via,
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
            destination_key = _canonical_graph_rank_key(
                sheet=None,
                rect=None,
                kind="opaque",
                label=via,
                via=via,
            )
        elif any(value is None for value in destination_values):
            raise _projection_corrupt(edge_id, "destination is a partial rectangle")
        else:
            destination_sheet_id = _projection_integral(
                row["dst_sheet_id"], edge_id, "destination sheet id"
            )
            if row["dst_sheet_name"] is None:
                raise _projection_corrupt(edge_id, "destination sheet does not exist")
            destination_sheet = str(row["dst_sheet_name"])
            if not destination_sheet:
                raise _projection_corrupt(edge_id, "destination sheet name is empty")
            destination_rect = _projection_rect(
                edge_id,
                row["dst_row_min"],
                row["dst_row_max"],
                row["dst_col_min"],
                row["dst_col_max"],
                "destination",
            )
            if (
                destination_rect.row_min == destination_rect.row_max
                and destination_rect.col_min == destination_rect.col_max
            ):
                destination_kind = "cell"
                destination_label = cell_symbol_id(
                    destination_sheet,
                    make_cell_ref(destination_rect.row_min, destination_rect.col_min),
                )
            else:
                destination_kind = "range"
                destination_label = _qualified_rect_ref(destination_sheet, destination_rect)
            destination_key = _canonical_graph_rank_key(
                sheet=destination_sheet,
                rect=destination_rect,
                kind=destination_kind,
                label=destination_label,
                via=via,
            )
        projections.append(
            (
                edge_id,
                source_sheet_id,
                source_rect,
                destination_sheet_id,
                destination_rect,
                source_key,
                destination_key,
                row["dependent_rank"],
                row["precedent_rank"],
            )
        )

    dependent_ranks = _canonical_dense_graph_ranks(item[5] for item in projections)
    precedent_ranks = _canonical_dense_graph_ranks(item[6] for item in projections)
    records = tuple(
        RankedEdge(
            edge_id=item[0],
            source_sheet_id=item[1],
            source_rect=item[2],
            destination_sheet_id=item[3],
            destination_rect=item[4],
            dependent_rank=dependent_ranks[item[5]],
            precedent_rank=precedent_ranks[item[6]],
            dependent_key=item[5],
            precedent_key=item[6],
        )
        for item in projections
    )
    if require_persisted_ranks:
        for item, record in zip(projections, records, strict=True):
            edge_id = record.edge_id
            dependent_rank = _projection_integral(item[7], edge_id, "persisted dependent rank")
            precedent_rank = _projection_integral(item[8], edge_id, "persisted precedent rank")
            if dependent_rank != record.dependent_rank or precedent_rank != record.precedent_rank:
                raise _projection_corrupt(
                    edge_id, "persisted ranks do not match canonical public-hop ordering"
                )
    return records


def _canonical_graph_rank_key(
    *,
    sheet: str | None,
    rect: Rect | None,
    kind: str,
    label: str,
    via: str,
) -> tuple[object, ...]:
    """Match the exact public GraphHop ordering used by graph queries."""
    if sheet is None or rect is None:
        return (
            1,
            "",
            "",
            0,
            0,
            0,
            0,
            kind,
            label.casefold(),
            label,
            via.casefold(),
            via,
        )
    return (
        0,
        sheet.casefold(),
        sheet,
        rect.row_min,
        rect.col_min,
        rect.row_max,
        rect.col_max,
        kind,
        label.casefold(),
        label,
        via.casefold(),
        via,
    )


def canonical_rank_key_text(key: tuple[object, ...]) -> str:
    """Serialize one canonical public-hop key without lossy coercion."""
    return json.dumps(key, ensure_ascii=False, separators=(",", ":"))


def _canonical_dense_graph_ranks(
    keys: Iterable[tuple[object, ...]],
) -> dict[tuple[object, ...], int]:
    ordered = sorted(set(keys))
    if len(ordered) > _GRAPH_RANK_MAX:
        raise GraphProjectionError(
            "dependency graph has too many unique public hops for the spatial index"
        )
    return {key: rank for rank, key in enumerate(ordered, start=1)}


def _projection_integral(value: object, edge_id: int, label: str) -> int:
    if type(value) is int:
        return value
    if type(value) is float and value.is_integer():
        return int(value)
    raise _projection_corrupt(edge_id, f"{label} is not an integer")


def _projection_rect(
    edge_id: int,
    row_min: object,
    row_max: object,
    col_min: object,
    col_max: object,
    label: str,
) -> Rect:
    coordinates = (
        _projection_integral(row_min, edge_id, f"{label} row minimum"),
        _projection_integral(row_max, edge_id, f"{label} row maximum"),
        _projection_integral(col_min, edge_id, f"{label} column minimum"),
        _projection_integral(col_max, edge_id, f"{label} column maximum"),
    )
    try:
        return Rect(*coordinates)
    except (TypeError, ValueError) as exc:
        raise _projection_corrupt(edge_id, f"{label} rectangle is invalid") from exc


def _projection_corrupt(edge_id: int, problem: str) -> GraphProjectionError:
    return GraphProjectionError(f"dependency edge {edge_id} is corrupt: {problem}")


def _qualified_rect_ref(sheet: str, rect: Rect) -> str:
    start = make_cell_ref(rect.row_min, rect.col_min)
    end = make_cell_ref(rect.row_max, rect.col_max)
    local_ref = start if start == end else f"{start}:{end}"
    if sheet.replace("_", "a").isalnum() and not sheet[0].isdigit():
        return f"{sheet}!{local_ref}"
    return f"'{sheet.replace(chr(39), chr(39) * 2)}'!{local_ref}"


@dataclass(frozen=True, slots=True)
class _IntervalEntry:
    edge_id: int
    sheet_id: int
    rect: Rect
    rank: int


@dataclass(frozen=True, slots=True)
class _IntervalNode:
    bounds: Rect
    min_rank: int
    max_rank: int
    entries: tuple[_IntervalEntry, ...] = ()
    left: _IntervalNode | None = None
    right: _IntervalNode | None = None


@dataclass(frozen=True, slots=True)
class _GraphTrustState:
    singleton: int
    dirty: int
    dependent_rank_max: int
    precedent_rank_max: int
    revision: int
    mutation_epoch: int
    clean_epoch: int


@dataclass(frozen=True, slots=True)
class _LiveGraphSeal:
    """Process-local monotonic evidence that validated graph storage stayed unchanged."""

    graph_write_epoch: int
    data_version: int
    schema_version: int
    temp_schema_version: int


def _release_initialization_transaction(
    connection: Any,
) -> tuple[BaseException, ...]:
    """Release an EdgeStore constructor lock, closing if release is unprovable."""
    errors: list[BaseException] = []
    try:
        connection.rollback()
    except BaseException as error:
        errors.append(error)

    native_connection = isinstance(connection, sqlite3.Connection)

    def transaction_active() -> bool:
        if native_connection:
            descriptor = cast(Any, sqlite3.Connection.in_transaction)
            return bool(descriptor.__get__(connection, sqlite3.Connection))
        return bool(connection.in_transaction)

    try:
        active = transaction_active()
    except BaseException as error:
        errors.append(error)
        active = True
    if active:
        try:
            if native_connection:
                sqlite3.Connection.execute(connection, "ROLLBACK")
            else:
                connection.execute("ROLLBACK")
        except BaseException as error:
            errors.append(error)
        try:
            active = transaction_active()
        except BaseException as error:
            errors.append(error)
            active = True
    if active:
        try:
            if native_connection:
                sqlite3.Connection.close(connection)
            else:
                connection.close()
        except BaseException as error:
            errors.append(error)
    unique: list[BaseException] = []
    seen: set[int] = set()
    for error in errors:
        if id(error) not in seen:
            seen.add(id(error))
            unique.append(error)
    return tuple(unique)


def _initialization_cleanup_failure(errors: tuple[BaseException, ...]) -> BaseException:
    failure: BaseException = (
        errors[0]
        if len(errors) == 1
        else BaseExceptionGroup("EdgeStore initialization cleanup failures", errors)
    )
    normalized = normalize_exception_graph(failure)
    if normalized is None:
        raise AssertionError("initialization cleanup evidence normalized to empty")
    return normalized


class EdgeStore:
    """Store/query edge destination rectangles without exposing the backend.

    R*Tree is preferred. SQLite builds without the module use a plain table and
    the same inclusive interval-overlap predicates. Keeping the fallback here
    prevents graph callers from acquiring backend-specific SQL.
    """

    RTREE_TABLE = "edge_rtree"
    INTERVAL_TABLE = "edge_intervals"
    SOURCE_RTREE_TABLE = "edge_source_rtree"
    SOURCE_INTERVAL_TABLE = "edge_source_intervals"
    MAX_PAGE_SIZE = 1_000

    def __init__(self, connection: SQLiteConnectionLike) -> None:
        self._connection = connection
        self._sealed_trust_state: _GraphTrustState | None = None
        self._pending_trust_state: _GraphTrustState | None = None
        self._transaction_seal_snapshot: _GraphTrustState | None = None
        self._sealed_catalog_keys: dict[tuple[str, int], str] | None = None
        self._pending_catalog_keys: dict[tuple[str, int], str] | None = None
        self._transaction_catalog_keys_snapshot: dict[tuple[str, int], str] | None = None
        self._sealed_live_seal: _LiveGraphSeal | None = None
        self._pending_live_seal: _LiveGraphSeal | None = None
        self._transaction_live_snapshot: _LiveGraphSeal | None = None
        self._transaction_active = False
        self._interval_cache: dict[EdgeDirection, dict[int, _IntervalNode | None]] | None = None
        self._interval_cache_revision: int | None = None
        if self._connection.in_transaction:
            raise RuntimeError("EdgeStore construction requires a connection outside a transaction")
        try:
            # Acquisition belongs to the cleanup scope because an instrumented
            # native connection can raise after BEGIN has taken effect.
            self._connection.execute("BEGIN IMMEDIATE")
            all_tables = self._all_tables()
            _reject_graph_temp_shadows(self._connection)
            self._graph_managed = "graph_spatial_state" in all_tables
            tables = all_tables & self._spatial_table_names()
            if tables == {self.RTREE_TABLE, self.SOURCE_RTREE_TABLE}:
                self.backend: EdgeBackend = "rtree"
            elif tables == {self.INTERVAL_TABLE, self.SOURCE_INTERVAL_TABLE}:
                self.backend = "interval"
            elif tables:
                raise EdgeSchemaError("edge range schema is partial or mixes spatial backends")
            else:
                raise EdgeSchemaError("edge range schema has not been initialized")
            _validate_graph_sidecar(self._connection, self.backend, all_tables)
            if "graph_spatial_state" in all_tables:
                self._sealed_trust_state = self._clean_trust_state()
                self._sealed_catalog_keys = self._catalog_keys()
                self._sealed_live_seal = self._live_graph_seal()
            if self.backend == "interval":
                self._warm_interval_cache_if_clean()
        except BaseException as primary_error:
            cleanup_errors = _release_initialization_transaction(self._connection)
            if cleanup_errors:
                primary_error.add_note(
                    "EdgeStore initialization cleanup also failed; "
                    "the original error remains primary."
                )
                cleanup_failure = prepare_chained_failure_with_primary_evidence(
                    _initialization_cleanup_failure(cleanup_errors),
                    primary_error,
                    message=("EdgeStore initialization causal evidence and cleanup failure"),
                )
                if cleanup_failure is not None:
                    raise primary_error from cleanup_failure
            raise
        else:
            cleanup_errors = _release_initialization_transaction(self._connection)
            if cleanup_errors:
                raise _initialization_cleanup_failure(cleanup_errors)

    def transaction_started(self) -> None:
        """Snapshot process-local trust for one outer IndexStore transaction."""
        if self._transaction_active:
            raise RuntimeError("edge-store transaction seal is already active")
        _reject_graph_temp_shadows(self._connection)
        self._transaction_active = True
        self._transaction_seal_snapshot = self._sealed_trust_state
        self._transaction_catalog_keys_snapshot = self._sealed_catalog_keys
        self._transaction_live_snapshot = self._sealed_live_seal
        self._pending_trust_state = None
        self._pending_catalog_keys = None
        self._pending_live_seal = None

    def transaction_committed(self) -> None:
        """Publish a successfully rebuilt epoch after the database commit."""
        self.finalize_transaction_commit()

    def finalize_transaction_commit(self) -> None:
        """Idempotently publish committed seal bookkeeping without SQLite I/O."""
        if not self._transaction_active:
            return
        if self._pending_trust_state is not None:
            self._sealed_trust_state = self._pending_trust_state
        if self._pending_catalog_keys is not None:
            self._sealed_catalog_keys = self._pending_catalog_keys
        live_seal = self._pending_live_seal or self._transaction_live_snapshot
        if live_seal is not None:
            # The graph-write epoch is a process-local monotonic mutation
            # identity.  Rebasing it here would bless writes performed after
            # the last validated rebuild (or any writes in a transaction that
            # did not rebuild).  A conservative stale facade is recoverable by
            # constructing a fresh EdgeStore; silently accepting a coherent
            # semantic forgery is not.
            self._sealed_live_seal = live_seal
        self._transaction_active = False
        self._transaction_seal_snapshot = None
        self._transaction_catalog_keys_snapshot = None
        self._transaction_live_snapshot = None
        self._pending_trust_state = None
        self._pending_catalog_keys = None
        self._pending_live_seal = None

    def transaction_rolled_back(self) -> None:
        """Restore the seal and interval cache after the database rollback."""
        was_active = self._transaction_active
        live_snapshot = self._transaction_live_snapshot
        self.finalize_transaction_rollback()
        if not was_active:
            return
        if self.backend == "rtree":
            # Rolling back TEMP DDL invalidates SQLite's connection-local schema
            # cache.  Reconnect both virtual tables before resealing: otherwise
            # their first later read replays virtual-table creation authorization
            # (without a write) and conservatively advances the tracked epoch.
            for table in (self.RTREE_TABLE, self.SOURCE_RTREE_TABLE):
                self._connection.execute(f"SELECT 1 FROM main.{table} LIMIT 1").fetchone()
            if live_snapshot is not None:
                self._sealed_live_seal = _LiveGraphSeal(
                    self._same_handle_graph_epoch(),
                    live_snapshot.data_version,
                    live_snapshot.schema_version,
                    live_snapshot.temp_schema_version,
                )
            return
        self._warm_interval_cache_if_clean()

    def finalize_transaction_rollback(self) -> None:
        """Idempotently restore rollback bookkeeping without SQLite I/O."""
        if not self._transaction_active:
            return
        self._sealed_trust_state = self._transaction_seal_snapshot
        self._sealed_catalog_keys = self._transaction_catalog_keys_snapshot
        if self._transaction_live_snapshot is not None:
            self._sealed_live_seal = _LiveGraphSeal(
                self._same_handle_graph_epoch(),
                self._transaction_live_snapshot.data_version,
                self._transaction_live_snapshot.schema_version,
                self._transaction_live_snapshot.temp_schema_version,
            )
        self._transaction_active = False
        self._transaction_seal_snapshot = None
        self._transaction_catalog_keys_snapshot = None
        self._transaction_live_snapshot = None
        self._pending_trust_state = None
        self._pending_catalog_keys = None
        self._pending_live_seal = None
        self._interval_cache = None
        self._interval_cache_revision = None

    def require_store_owned_transaction(self) -> None:
        """Reject graph mutation inside a transaction not opened by IndexStore."""
        _reject_graph_temp_shadows(self._connection)
        if self._graph_managed and self._connection.in_transaction and not self._transaction_active:
            raise RuntimeError(
                "Cannot adopt a raw SQLite transaction for graph mutation; "
                "use `with store.transaction():` to open the transaction."
            )

    def require_no_temp_shadows(self) -> None:
        """Reject connection-local objects that can shadow authoritative graph data."""
        _reject_graph_temp_shadows(self._connection)

    @classmethod
    def ensure_schema(
        cls, connection: SQLiteConnectionLike, *, prefer_rtree: bool = True
    ) -> EdgeBackend:
        """Create one spatial backend and return the selected implementation."""
        _reject_graph_temp_shadows(connection)
        tables = {
            str(row[0])
            for row in connection.execute(
                "SELECT name FROM main.sqlite_master WHERE type = 'table'"
            ).fetchall()
        }
        spatial_tables = tables & cls._spatial_table_names()
        if spatial_tables == {cls.RTREE_TABLE, cls.SOURCE_RTREE_TABLE}:
            _validate_graph_sidecar(connection, "rtree", tables)
            return "rtree"
        if spatial_tables == {cls.INTERVAL_TABLE, cls.SOURCE_INTERVAL_TABLE}:
            _validate_graph_sidecar(connection, "interval", tables)
            return "interval"
        if spatial_tables:
            raise EdgeSchemaError("edge range schema is partial or mixes spatial backends")

        if prefer_rtree:
            try:
                connection.execute(
                    "CREATE VIRTUAL TABLE edge_rtree USING rtree_i32("
                    "edge_id, sheet_min, sheet_max, row_min, row_max, col_min, col_max, "
                    "rank_min, rank_max)"
                )
                connection.execute(
                    "CREATE VIRTUAL TABLE edge_source_rtree USING rtree_i32("
                    "edge_id, sheet_min, sheet_max, row_min, row_max, col_min, col_max, "
                    "rank_min, rank_max)"
                )
                tables.update((cls.RTREE_TABLE, cls.SOURCE_RTREE_TABLE))
                _ensure_spatial_dirty_triggers(connection, "rtree", tables)
                _validate_graph_sidecar(connection, "rtree", tables)
                return "rtree"
            except sqlite3.OperationalError as exc:
                # Only module availability is recoverable. Syntax, disk, and
                # corruption failures must not be silently hidden.
                message = str(exc).casefold()
                if "no such module" not in message or "rtree" not in message:
                    raise

        connection.execute(
            """
            CREATE TABLE main.edge_intervals (
                edge_id INTEGER PRIMARY KEY,
                sheet_id INTEGER NOT NULL,
                row_min INTEGER NOT NULL,
                row_max INTEGER NOT NULL,
                col_min INTEGER NOT NULL,
                col_max INTEGER NOT NULL,
                rank INTEGER NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE INDEX main.edge_intervals_overlap
            ON edge_intervals(sheet_id, rank, row_min, row_max, col_min, col_max)
            """
        )
        connection.execute(
            """
            CREATE TABLE main.edge_source_intervals (
                edge_id INTEGER PRIMARY KEY,
                sheet_id INTEGER NOT NULL,
                row_min INTEGER NOT NULL,
                row_max INTEGER NOT NULL,
                col_min INTEGER NOT NULL,
                col_max INTEGER NOT NULL,
                rank INTEGER NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE INDEX main.edge_source_intervals_overlap
            ON edge_source_intervals(sheet_id, rank, row_min, row_max, col_min, col_max)
            """
        )
        tables.update((cls.INTERVAL_TABLE, cls.SOURCE_INTERVAL_TABLE))
        _ensure_spatial_dirty_triggers(connection, "interval", tables)
        _validate_graph_sidecar(connection, "interval", tables)
        return "interval"

    def clear(self) -> None:
        """Remove both derived spatial mirrors."""
        self._connection.execute(f"DELETE FROM main.{self.table_name}")
        self._connection.execute(f"DELETE FROM main.{self.source_table_name}")
        self._mark_dirty()

    def insert(self, edge_id: int, sheet_id: int, rect: Rect) -> None:
        """Insert or replace one inclusive destination rectangle."""
        if self.backend == "rtree":
            self._connection.execute(
                "INSERT OR REPLACE INTO main.edge_rtree("
                "edge_id, sheet_min, sheet_max, row_min, row_max, col_min, col_max, "
                "rank_min, rank_max) VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0)",
                (
                    edge_id,
                    sheet_id,
                    sheet_id,
                    rect.row_min,
                    rect.row_max,
                    rect.col_min,
                    rect.col_max,
                ),
            )
            self._mark_dirty()
            return
        self._connection.execute(
            "INSERT OR REPLACE INTO main.edge_intervals("
            "edge_id, sheet_id, row_min, row_max, col_min, col_max, rank"
            ") VALUES (?, ?, ?, ?, ?, ?, 0)",
            (
                edge_id,
                sheet_id,
                rect.row_min,
                rect.row_max,
                rect.col_min,
                rect.col_max,
            ),
        )
        self._mark_dirty()

    def delete(self, edge_id: int) -> None:
        """Delete one edge from both spatial mirrors if present."""
        self._connection.execute(
            f"DELETE FROM main.{self.table_name} WHERE edge_id = ?", (edge_id,)
        )
        self._connection.execute(
            f"DELETE FROM main.{self.source_table_name} WHERE edge_id = ?", (edge_id,)
        )
        self._mark_dirty()

    def rebuild_ranked(self, records: Iterable[RankedEdge]) -> None:
        """Atomically replace both mirrors from globally ranked edge records."""
        self.require_store_owned_transaction()
        items = tuple(records)
        _validate_ranked_records(items)
        relational_ids = {
            int(row[0]) for row in self._connection.execute("SELECT id FROM main.edges")
        }
        item_ids = {item.edge_id for item in items}
        if item_ids != relational_ids:
            raise RuntimeError("ranked records do not cover dependency edges exactly")
        self.clear()
        self._connection.executemany(
            "UPDATE main.edges SET dependent_rank = ?, precedent_rank = ? WHERE id = ?",
            ((item.dependent_rank, item.precedent_rank, item.edge_id) for item in items),
        )
        if self.backend == "rtree":
            self._connection.executemany(
                """
                INSERT INTO main.edge_source_rtree(
                    edge_id, sheet_min, sheet_max, row_min, row_max,
                    col_min, col_max, rank_min, rank_max
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    (
                        item.edge_id,
                        item.source_sheet_id,
                        item.source_sheet_id,
                        item.source_rect.row_min,
                        item.source_rect.row_max,
                        item.source_rect.col_min,
                        item.source_rect.col_max,
                        item.precedent_rank,
                        item.precedent_rank,
                    )
                    for item in items
                ),
            )
            self._connection.executemany(
                """
                INSERT INTO main.edge_rtree(
                    edge_id, sheet_min, sheet_max, row_min, row_max,
                    col_min, col_max, rank_min, rank_max
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    (
                        item.edge_id,
                        item.destination_sheet_id,
                        item.destination_sheet_id,
                        item.destination_rect.row_min,
                        item.destination_rect.row_max,
                        item.destination_rect.col_min,
                        item.destination_rect.col_max,
                        item.dependent_rank,
                        item.dependent_rank,
                    )
                    for item in items
                    if item.destination_sheet_id is not None and item.destination_rect is not None
                ),
            )
        else:
            self._connection.executemany(
                """
                INSERT INTO main.edge_source_intervals(
                    edge_id, sheet_id, row_min, row_max, col_min, col_max, rank
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    (
                        item.edge_id,
                        item.source_sheet_id,
                        item.source_rect.row_min,
                        item.source_rect.row_max,
                        item.source_rect.col_min,
                        item.source_rect.col_max,
                        item.precedent_rank,
                    )
                    for item in items
                ),
            )
            self._connection.executemany(
                """
                INSERT INTO main.edge_intervals(
                    edge_id, sheet_id, row_min, row_max, col_min, col_max, rank
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    (
                        item.edge_id,
                        item.destination_sheet_id,
                        item.destination_rect.row_min,
                        item.destination_rect.row_max,
                        item.destination_rect.col_min,
                        item.destination_rect.col_max,
                        item.dependent_rank,
                    )
                    for item in items
                    if item.destination_sheet_id is not None and item.destination_rect is not None
                ),
            )
        destination_count = sum(item.destination_rect is not None for item in items)
        if self._count(self.source_table_name) != len(items):
            raise RuntimeError("source spatial mirror count does not match dependency edges")
        if self._count(self.table_name) != destination_count:
            raise RuntimeError("destination spatial mirror count does not match range edges")
        dependent_rank_max = max((item.dependent_rank for item in items), default=0)
        precedent_rank_max = max((item.precedent_rank for item in items), default=0)
        rank_keys: dict[tuple[str, int], str] = {}
        for item in items:
            rank_keys[("dependents", item.dependent_rank)] = canonical_rank_key_text(
                item.dependent_key
            )
            rank_keys[("precedents", item.precedent_rank)] = canonical_rank_key_text(
                item.precedent_key
            )
        # Populate only after every active mirror mutation trigger has run.
        self._connection.execute("DELETE FROM main.graph_rank_keys")
        self._connection.executemany(
            "INSERT INTO main.graph_rank_keys(direction, rank, key_text) VALUES (?, ?, ?)",
            (
                (direction, rank, key_text)
                for (direction, rank), key_text in sorted(rank_keys.items())
            ),
        )
        self._connection.execute(
            """
            UPDATE main.graph_spatial_state
            SET dirty = 0, dependent_rank_max = ?, precedent_rank_max = ?,
                revision = revision + 1, clean_epoch = mutation_epoch
            WHERE singleton = 1
            """,
            (dependent_rank_max, precedent_rank_max),
        )
        if self.backend == "interval":
            self._refresh_interval_cache()
        self._seal_successful_rebuild()

    def require_clean(self) -> _GraphTrustState:
        """Reject graph queries while ranked mirrors are stale or absent."""
        current_state = self._clean_trust_state()
        if current_state is None:
            raise RuntimeError("ranked dependency mirrors are dirty")
        expected_state = (
            self._pending_trust_state
            if self._transaction_active and self._pending_trust_state is not None
            else self._sealed_trust_state
        )
        if expected_state is None or current_state != expected_state:
            raise RuntimeError("ranked dependency mirror seal does not match persisted state")
        expected_live_seal = (
            self._pending_live_seal
            if self._transaction_active and self._pending_live_seal is not None
            else self._sealed_live_seal
        )
        if expected_live_seal is None or self._live_graph_seal() != expected_live_seal:
            raise RuntimeError("ranked dependency live seal changed after validation")
        return current_state

    def max_rank(self, direction: EdgeDirection) -> int:
        """Return the greatest dense public-hop rank for one direction."""
        state = self.require_clean()
        if direction == "dependents":
            return state.dependent_rank_max
        if direction == "precedents":
            return state.precedent_rank_max
        raise ValueError(f"unsupported edge direction: {direction}")

    def first_matching_rank(
        self,
        direction: EdgeDirection,
        sheet_id: int,
        rect: Rect,
        *,
        after_rank: int = 0,
    ) -> int | None:
        """Find the first semantic rank intersecting ``rect`` without sorting hits."""
        self.require_clean()
        _validate_rank_query(sheet_id, rect, after_rank)
        maximum = self.max_rank(direction)
        if after_rank >= maximum:
            return None
        table = self._direction_table(direction)
        if self.backend == "interval":
            del table
            return self._interval_first_matching_rank(direction, sheet_id, rect, after_rank)

        lower = after_rank + 1
        upper = maximum
        if not self._rtree_rank_exists(table, sheet_id, rect, lower, upper):
            return None
        while lower < upper:
            middle = lower + (upper - lower) // 2
            if self._rtree_rank_exists(table, sheet_id, rect, lower, middle):
                upper = middle
            else:
                lower = middle + 1
        return lower

    def edge_id_at_rank(
        self,
        direction: EdgeDirection,
        sheet_id: int,
        rect: Rect,
        rank: int,
    ) -> int | None:
        """Return one representative matching edge for a unique public-hop rank."""
        self.require_clean()
        if type(rank) is not int or rank < 1:
            raise ValueError("rank must be a positive integer")
        _validate_rank_query(sheet_id, rect, rank - 1)
        table = self._direction_table(direction)
        rank_column = "rank_min" if self.backend == "rtree" else "rank"
        sheet_predicate = (
            "sheet_min <= ? AND sheet_max >= ?" if self.backend == "rtree" else "sheet_id = ?"
        )
        sheet_parameters: tuple[int, ...] = (
            (sheet_id, sheet_id) if self.backend == "rtree" else (sheet_id,)
        )
        row = self._connection.execute(
            f"""
            SELECT edge_id FROM main.{table}
            WHERE {sheet_predicate} AND {rank_column} = ?
              AND row_min <= ? AND row_max >= ?
              AND col_min <= ? AND col_max >= ?
            LIMIT 1
            """,
            (
                *sheet_parameters,
                rank,
                rect.row_max,
                rect.row_min,
                rect.col_max,
                rect.col_min,
            ),
        ).fetchone()
        return None if row is None else _exact_spatial_int(row[0], "edge id")

    def rank_key_text(self, direction: EdgeDirection, rank: int) -> str:
        """Return the immutable canonical identity for one selected dense rank."""
        self.require_clean()
        if direction not in ("dependents", "precedents"):
            raise ValueError(f"unsupported edge direction: {direction}")
        if type(rank) is not int or rank < 1:
            raise ValueError("rank must be a positive integer")
        row = self._connection.execute(
            "SELECT key_text FROM main.graph_rank_keys WHERE direction = ? AND rank = ?",
            (direction, rank),
        ).fetchone()
        if row is None or type(row[0]) is not str or not row[0]:
            raise RuntimeError("ranked dependency semantic identity is missing or corrupt")
        expected_keys = (
            self._pending_catalog_keys
            if self._transaction_active and self._pending_catalog_keys is not None
            else self._sealed_catalog_keys
        )
        if expected_keys is None or expected_keys.get((direction, rank)) != row[0]:
            raise RuntimeError("ranked dependency semantic identity changed after validation")
        return row[0]

    def destination(self, edge_id: int) -> tuple[int, Rect] | None:
        """Return one exact spatial mirror entry, if present."""
        if type(edge_id) is not int or edge_id < 1:
            raise ValueError("edge_id must be a positive integer")
        if self.backend == "rtree":
            row = self._connection.execute(
                """
                SELECT sheet_min, sheet_max, row_min, row_max, col_min, col_max
                FROM main.edge_rtree WHERE edge_id = ?
                """,
                (edge_id,),
            ).fetchone()
            if row is None:
                return None
            sheet_min = _exact_spatial_int(row[0], "sheet minimum")
            sheet_max = _exact_spatial_int(row[1], "sheet maximum")
            if sheet_min != sheet_max:
                raise ValueError("spatial destination has an invalid sheet dimension")
            return (
                sheet_min,
                Rect(
                    _exact_spatial_int(row[2], "row minimum"),
                    _exact_spatial_int(row[3], "row maximum"),
                    _exact_spatial_int(row[4], "column minimum"),
                    _exact_spatial_int(row[5], "column maximum"),
                ),
            )
        row = self._connection.execute(
            """
            SELECT sheet_id, row_min, row_max, col_min, col_max
            FROM main.edge_intervals WHERE edge_id = ?
            """,
            (edge_id,),
        ).fetchone()
        if row is None:
            return None
        return (
            _exact_spatial_int(row[0], "sheet id"),
            Rect(
                _exact_spatial_int(row[1], "row minimum"),
                _exact_spatial_int(row[2], "row maximum"),
                _exact_spatial_int(row[3], "column minimum"),
                _exact_spatial_int(row[4], "column maximum"),
            ),
        )

    def ranked_mirror(self, edge_id: int, direction: EdgeDirection) -> tuple[int, Rect, int] | None:
        """Return one exact ranked mirror row for corruption validation."""
        if type(edge_id) is not int or edge_id < 1:
            raise ValueError("edge_id must be a positive integer")
        table = self._direction_table(direction)
        if self.backend == "rtree":
            row = self._connection.execute(
                f"""
                SELECT sheet_min, sheet_max, row_min, row_max,
                       col_min, col_max, rank_min, rank_max
                FROM main.{table} WHERE edge_id = ?
                """,
                (edge_id,),
            ).fetchone()
            if row is None:
                return None
            sheet = _equal_spatial_pair(row[0], row[1], "sheet")
            rank = _equal_spatial_pair(row[6], row[7], "rank")
            return (
                sheet,
                Rect(
                    _exact_spatial_int(row[2], "row minimum"),
                    _exact_spatial_int(row[3], "row maximum"),
                    _exact_spatial_int(row[4], "column minimum"),
                    _exact_spatial_int(row[5], "column maximum"),
                ),
                rank,
            )
        row = self._connection.execute(
            f"""
            SELECT sheet_id, row_min, row_max, col_min, col_max, rank
            FROM main.{table} WHERE edge_id = ?
            """,
            (edge_id,),
        ).fetchone()
        if row is None:
            return None
        return (
            _exact_spatial_int(row[0], "sheet id"),
            Rect(
                _exact_spatial_int(row[1], "row minimum"),
                _exact_spatial_int(row[2], "row maximum"),
                _exact_spatial_int(row[3], "column minimum"),
                _exact_spatial_int(row[4], "column maximum"),
            ),
            _exact_spatial_int(row[5], "rank"),
        )

    def query_point(self, sheet_id: int, row: int, col: int) -> tuple[int, ...]:
        """Return edge ids whose destinations contain a worksheet point."""
        return tuple(self.iter_query_point(sheet_id, row, col))

    def query_point_page(
        self,
        sheet_id: int,
        row: int,
        col: int,
        *,
        after_edge_id: int = 0,
        limit: int = 256,
    ) -> tuple[int, ...]:
        """Return one edge-id keyset page for a worksheet point."""
        _validate_query(sheet_id, Rect(row, row, col, col), after_edge_id, limit)
        if self.backend == "rtree":
            rows = self._connection.execute(
                """
                SELECT edge_id FROM main.edge_rtree
                WHERE sheet_min <= ? AND sheet_max >= ?
                  AND row_min <= ? AND row_max >= ?
                  AND col_min <= ? AND col_max >= ?
                  AND edge_id > ?
                ORDER BY edge_id
                LIMIT ?
                """,
                (sheet_id, sheet_id, row, row, col, col, after_edge_id, limit),
            )
        else:
            rows = self._connection.execute(
                """
                SELECT edge_id FROM main.edge_intervals
                WHERE sheet_id = ?
                  AND row_min <= ? AND row_max >= ?
                  AND col_min <= ? AND col_max >= ?
                  AND edge_id > ?
                ORDER BY edge_id
                LIMIT ?
                """,
                (sheet_id, row, row, col, col, after_edge_id, limit),
            )
        return tuple(int(item[0]) for item in rows.fetchall())

    def iter_query_point(
        self, sheet_id: int, row: int, col: int, *, page_size: int = 256
    ) -> Iterator[int]:
        """Iterate matching edge ids with bounded keyset pages."""
        after_edge_id = 0
        while page := self.query_point_page(
            sheet_id, row, col, after_edge_id=after_edge_id, limit=page_size
        ):
            yield from page
            after_edge_id = page[-1]

    def query_range(self, sheet_id: int, rect: Rect) -> tuple[int, ...]:
        """Return edge ids whose destinations overlap an inclusive rectangle."""
        return tuple(self.iter_query_range(sheet_id, rect))

    def query_range_page(
        self,
        sheet_id: int,
        rect: Rect,
        *,
        after_edge_id: int = 0,
        limit: int = 256,
    ) -> tuple[int, ...]:
        """Return one edge-id keyset page for an inclusive rectangle."""
        _validate_query(sheet_id, rect, after_edge_id, limit)
        if self.backend == "rtree":
            rows = self._connection.execute(
                """
                SELECT edge_id FROM main.edge_rtree
                WHERE sheet_min <= ? AND sheet_max >= ?
                  AND row_min <= ? AND row_max >= ?
                  AND col_min <= ? AND col_max >= ?
                  AND edge_id > ?
                ORDER BY edge_id
                LIMIT ?
                """,
                (
                    sheet_id,
                    sheet_id,
                    rect.row_max,
                    rect.row_min,
                    rect.col_max,
                    rect.col_min,
                    after_edge_id,
                    limit,
                ),
            )
        else:
            rows = self._connection.execute(
                """
                SELECT edge_id FROM main.edge_intervals
                WHERE sheet_id = ?
                  AND row_min <= ? AND row_max >= ?
                  AND col_min <= ? AND col_max >= ?
                  AND edge_id > ?
                ORDER BY edge_id
                LIMIT ?
                """,
                (
                    sheet_id,
                    rect.row_max,
                    rect.row_min,
                    rect.col_max,
                    rect.col_min,
                    after_edge_id,
                    limit,
                ),
            )
        return tuple(int(item[0]) for item in rows.fetchall())

    def iter_query_range(self, sheet_id: int, rect: Rect, *, page_size: int = 256) -> Iterator[int]:
        """Iterate overlapping edge ids with bounded keyset pages."""
        after_edge_id = 0
        while page := self.query_range_page(
            sheet_id, rect, after_edge_id=after_edge_id, limit=page_size
        ):
            yield from page
            after_edge_id = page[-1]

    @property
    def table_name(self) -> str:
        """Return the physical table used by this connection."""
        return self.RTREE_TABLE if self.backend == "rtree" else self.INTERVAL_TABLE

    @property
    def source_table_name(self) -> str:
        """Return the physical source-rectangle table for this connection."""
        return self.SOURCE_RTREE_TABLE if self.backend == "rtree" else self.SOURCE_INTERVAL_TABLE

    def _direction_table(self, direction: EdgeDirection) -> str:
        if direction == "dependents":
            return self.table_name
        if direction == "precedents":
            return self.source_table_name
        raise ValueError("direction must be 'dependents' or 'precedents'")

    def _rtree_rank_exists(
        self,
        table: str,
        sheet_id: int,
        rect: Rect,
        rank_min: int,
        rank_max: int,
    ) -> bool:
        return (
            self._connection.execute(
                f"""
                SELECT 1 FROM main.{table}
                WHERE sheet_min <= ? AND sheet_max >= ?
                  AND row_min <= ? AND row_max >= ?
                  AND col_min <= ? AND col_max >= ?
                  AND rank_min >= ? AND rank_max <= ?
                LIMIT 1
                """,
                (
                    sheet_id,
                    sheet_id,
                    rect.row_max,
                    rect.row_min,
                    rect.col_max,
                    rect.col_min,
                    rank_min,
                    rank_max,
                ),
            ).fetchone()
            is not None
        )

    def _clean_trust_state(self) -> _GraphTrustState | None:
        row = self._connection.execute(
            """
            SELECT singleton, dirty, dependent_rank_max, precedent_rank_max,
                   revision, mutation_epoch, clean_epoch
            FROM main.graph_spatial_state WHERE singleton = 1
            """
        ).fetchone()
        if row is None:
            return None
        state = _GraphTrustState(*(_exact_spatial_int(value, "graph trust state") for value in row))
        if (
            state.singleton != 1
            or state.dirty != 0
            or state.dependent_rank_max < 0
            or state.precedent_rank_max < 0
            or state.revision < 0
            or state.mutation_epoch < 0
            or state.clean_epoch < 0
            or state.mutation_epoch != state.clean_epoch
        ):
            return None
        return state

    def _seal_successful_rebuild(self) -> None:
        state = self._clean_trust_state()
        if state is None:
            raise RuntimeError("successful graph rebuild did not leave clean trusted state")
        catalog_keys = self._catalog_keys()
        live_seal = self._live_graph_seal()
        if self._transaction_active:
            self._pending_trust_state = state
            self._pending_catalog_keys = catalog_keys
            self._pending_live_seal = live_seal
        else:
            self._sealed_trust_state = state
            self._sealed_catalog_keys = catalog_keys
            self._sealed_live_seal = live_seal

    def _catalog_keys(self) -> dict[tuple[str, int], str]:
        return {
            (str(row[0]), _exact_spatial_int(row[1], "graph rank catalog rank")): str(row[2])
            for row in self._connection.execute(
                "SELECT direction, rank, key_text FROM main.graph_rank_keys "
                "ORDER BY direction, rank"
            )
        }

    def _live_graph_seal(self) -> _LiveGraphSeal:
        """Capture monotonic same-handle, other-handle, and DDL identities."""
        data_version = self._pragma_int("data_version")
        schema_version = self._pragma_int("schema_version")
        temp_schema_version = self._pragma_int("schema_version", database="temp")
        return _LiveGraphSeal(
            self._same_handle_graph_epoch(),
            data_version,
            schema_version,
            temp_schema_version,
        )

    def _same_handle_graph_epoch(self) -> int:
        """Use graph-specific tracking when available, conservatively otherwise."""
        value = getattr(self._connection, "_graph_write_epoch", None)
        if value is None:
            value = self._connection.total_changes
        if type(value) is not int or value < 0:
            raise RuntimeError("SQLite graph-write epoch is invalid")
        return value

    def _pragma_int(
        self,
        name: Literal["data_version", "schema_version"],
        *,
        database: Literal["main", "temp"] = "main",
    ) -> int:
        row = self._connection.execute(f"PRAGMA {database}.{name}").fetchone()
        if row is None:
            raise RuntimeError(f"SQLite {name.replace('_', ' ')} is unavailable")
        value = _exact_spatial_int(row[0], f"SQLite {name.replace('_', ' ')}")
        if value < 0:
            raise RuntimeError(f"SQLite {name.replace('_', ' ')} is negative")
        return value

    def _warm_interval_cache_if_clean(self) -> None:
        if not self._table_exists("graph_spatial_state"):
            return
        row = self._connection.execute(
            """
            SELECT dirty, revision, mutation_epoch, clean_epoch
            FROM main.graph_spatial_state WHERE singleton = 1
            """
        ).fetchone()
        if (
            row is not None
            and _exact_spatial_int(row[0], "dirty flag") == 0
            and _exact_spatial_int(row[2], "mutation epoch")
            == _exact_spatial_int(row[3], "clean epoch")
        ):
            self._refresh_interval_cache()

    def _refresh_interval_cache(self) -> None:
        if self.backend != "interval":
            return
        state = self._connection.execute(
            """
            SELECT dirty, revision, mutation_epoch, clean_epoch
            FROM main.graph_spatial_state WHERE singleton = 1
            """
        ).fetchone()
        if (
            state is None
            or _exact_spatial_int(state[0], "dirty flag") != 0
            or _exact_spatial_int(state[2], "mutation epoch")
            != _exact_spatial_int(state[3], "clean epoch")
        ):
            self._interval_cache = None
            self._interval_cache_revision = None
            return
        cache: dict[EdgeDirection, dict[int, _IntervalNode | None]] = {
            "dependents": {},
            "precedents": {},
        }
        for direction in ("dependents", "precedents"):
            by_sheet: dict[int, list[_IntervalEntry]] = {}
            for row in self._connection.execute(
                f"""
                SELECT edge_id, sheet_id, row_min, row_max, col_min, col_max, rank
                FROM main.{self._direction_table(direction)}
                ORDER BY sheet_id, rank, edge_id
                """
            ):
                edge_id = _exact_spatial_int(row[0], "edge id")
                sheet_id = _exact_spatial_int(row[1], "sheet id")
                entry = _IntervalEntry(
                    edge_id,
                    sheet_id,
                    Rect(
                        _exact_spatial_int(row[2], "row minimum"),
                        _exact_spatial_int(row[3], "row maximum"),
                        _exact_spatial_int(row[4], "column minimum"),
                        _exact_spatial_int(row[5], "column maximum"),
                    ),
                    _exact_spatial_int(row[6], "rank"),
                )
                by_sheet.setdefault(sheet_id, []).append(entry)
            cache[direction] = {
                sheet_id: _build_interval_tree(tuple(entries))
                for sheet_id, entries in sorted(by_sheet.items())
            }
        self._interval_cache = cache
        self._interval_cache_revision = _exact_spatial_int(state[1], "revision")

    def _interval_first_matching_rank(
        self,
        direction: EdgeDirection,
        sheet_id: int,
        rect: Rect,
        after_rank: int,
    ) -> int | None:
        state = self._connection.execute(
            """
            SELECT dirty, revision, mutation_epoch, clean_epoch
            FROM main.graph_spatial_state WHERE singleton = 1
            """
        ).fetchone()
        if (
            state is None
            or _exact_spatial_int(state[0], "dirty flag") != 0
            or _exact_spatial_int(state[2], "mutation epoch")
            != _exact_spatial_int(state[3], "clean epoch")
        ):
            raise RuntimeError("ranked dependency mirrors are dirty")
        revision = _exact_spatial_int(state[1], "revision")
        if self._interval_cache is None or self._interval_cache_revision != revision:
            self._refresh_interval_cache()
        if self._interval_cache is None:
            raise RuntimeError("ranked interval cache is unavailable")
        node = self._interval_cache[direction].get(sheet_id)
        return _interval_tree_min_rank(node, rect, after_rank)

    def _count(self, table: str) -> int:
        row = self._connection.execute(f"SELECT COUNT(*) FROM main.{table}").fetchone()
        if row is None:
            raise RuntimeError(f"could not count {table}")
        return int(row[0])

    def _mark_dirty(self) -> None:
        self._interval_cache = None
        self._interval_cache_revision = None
        if self._table_exists("graph_spatial_state"):
            self._connection.execute(
                """
                UPDATE main.graph_spatial_state
                SET dirty = 1, mutation_epoch = mutation_epoch + 1
                WHERE singleton = 1
                """
            )

    def _table_exists(self, table_name: str) -> bool:
        row = self._connection.execute(
            "SELECT 1 FROM main.sqlite_master WHERE type = 'table' AND name = ?",
            (table_name,),
        ).fetchone()
        return row is not None

    def _all_tables(self) -> set[str]:
        return {
            str(row[0])
            for row in self._connection.execute(
                "SELECT name FROM main.sqlite_master WHERE type = 'table'"
            )
        }

    @classmethod
    def _spatial_table_names(cls) -> set[str]:
        return {
            cls.RTREE_TABLE,
            cls.SOURCE_RTREE_TABLE,
            cls.INTERVAL_TABLE,
            cls.SOURCE_INTERVAL_TABLE,
        }


def _validate_query(sheet_id: int, rect: Rect, after_edge_id: int, limit: int) -> None:
    if type(sheet_id) is not int or sheet_id < 1:
        raise ValueError("sheet_id must be a positive integer")
    if type(after_edge_id) is not int or after_edge_id < 0:
        raise ValueError("after_edge_id must be a nonnegative integer")
    if type(limit) is not int or not 1 <= limit <= EdgeStore.MAX_PAGE_SIZE:
        raise ValueError(f"limit must be between 1 and {EdgeStore.MAX_PAGE_SIZE}")
    del rect


def _validate_rank_query(sheet_id: int, rect: Rect, after_rank: int) -> None:
    if type(sheet_id) is not int or sheet_id < 1:
        raise ValueError("sheet_id must be a positive integer")
    if type(after_rank) is not int or after_rank < 0:
        raise ValueError("rank must be a positive integer")
    del rect


def _build_interval_tree(entries: tuple[_IntervalEntry, ...]) -> _IntervalNode | None:
    if not entries:
        return None
    bounds = Rect(
        min(entry.rect.row_min for entry in entries),
        max(entry.rect.row_max for entry in entries),
        min(entry.rect.col_min for entry in entries),
        max(entry.rect.col_max for entry in entries),
    )
    min_rank = min(entry.rank for entry in entries)
    max_rank = max(entry.rank for entry in entries)
    if len(entries) <= 8:
        return _IntervalNode(
            bounds,
            min_rank,
            max_rank,
            tuple(sorted(entries, key=_interval_entry_key)),
        )
    row_centers = tuple(entry.rect.row_min + entry.rect.row_max for entry in entries)
    col_centers = tuple(entry.rect.col_min + entry.rect.col_max for entry in entries)
    split_rows = max(row_centers) - min(row_centers) >= max(col_centers) - min(col_centers)
    ordered = tuple(
        sorted(
            entries,
            key=lambda entry: (
                entry.rect.row_min + entry.rect.row_max
                if split_rows
                else entry.rect.col_min + entry.rect.col_max,
                *_interval_entry_key(entry),
            ),
        )
    )
    midpoint = len(ordered) // 2
    return _IntervalNode(
        bounds,
        min_rank,
        max_rank,
        left=_build_interval_tree(ordered[:midpoint]),
        right=_build_interval_tree(ordered[midpoint:]),
    )


def _interval_tree_min_rank(
    node: _IntervalNode | None,
    query: Rect,
    after_rank: int,
    best: int | None = None,
) -> int | None:
    if (
        node is None
        or node.max_rank <= after_rank
        or (best is not None and node.min_rank >= best)
        or not node.bounds.intersects(query)
    ):
        return best
    if node.entries:
        for entry in node.entries:
            if entry.rank <= after_rank or (best is not None and entry.rank >= best):
                continue
            if entry.rect.intersects(query):
                best = entry.rank
        return best
    children = tuple(
        sorted(
            (child for child in (node.left, node.right) if child is not None),
            key=lambda child: (child.min_rank, child.bounds.row_min, child.bounds.col_min),
        )
    )
    for child in children:
        best = _interval_tree_min_rank(child, query, after_rank, best)
    return best


def _interval_entry_key(entry: _IntervalEntry) -> tuple[int, int, int, int, int, int]:
    return (
        entry.rank,
        entry.rect.row_min,
        entry.rect.col_min,
        entry.rect.row_max,
        entry.rect.col_max,
        entry.edge_id,
    )


def _validate_ranked_records(records: tuple[RankedEdge, ...]) -> None:
    edge_ids: set[int] = set()
    dependent_ranks: set[int] = set()
    precedent_ranks: set[int] = set()
    dependent_key_to_rank: dict[tuple[object, ...], int] = {}
    dependent_rank_to_key: dict[int, tuple[object, ...]] = {}
    precedent_key_to_rank: dict[tuple[object, ...], int] = {}
    precedent_rank_to_key: dict[int, tuple[object, ...]] = {}
    for item in records:
        if type(item.edge_id) is not int or item.edge_id < 1 or item.edge_id in edge_ids:
            raise ValueError("ranked edge ids must be unique positive integers")
        edge_ids.add(item.edge_id)
        if (item.destination_sheet_id is None) != (item.destination_rect is None):
            raise ValueError(
                "destination sheet and rectangle must either both be set or both be null"
            )
        for label, rank in (
            ("dependent", item.dependent_rank),
            ("precedent", item.precedent_rank),
        ):
            if type(rank) is not int or not 1 <= rank <= 2_147_483_647:
                raise ValueError(f"{label} rank is outside the signed 32-bit range")
        dependent_ranks.add(item.dependent_rank)
        precedent_ranks.add(item.precedent_rank)
        _record_rank_identity(
            "dependent",
            item.dependent_key,
            item.dependent_rank,
            dependent_key_to_rank,
            dependent_rank_to_key,
        )
        _record_rank_identity(
            "precedent",
            item.precedent_key,
            item.precedent_rank,
            precedent_key_to_rank,
            precedent_rank_to_key,
        )
    for label, ranks in (("dependent", dependent_ranks), ("precedent", precedent_ranks)):
        if ranks and len(ranks) != max(ranks):
            raise ValueError(f"{label} ranks must be dense from one")


def _record_rank_identity(
    label: str,
    key: tuple[object, ...],
    rank: int,
    key_to_rank: dict[tuple[object, ...], int],
    rank_to_key: dict[int, tuple[object, ...]],
) -> None:
    if type(key) is not tuple or not key:
        raise ValueError(f"{label} semantic key must be a nonempty tuple")
    try:
        existing_rank = key_to_rank.setdefault(key, rank)
    except TypeError as exc:
        raise ValueError(f"{label} semantic key must be hashable") from exc
    if existing_rank != rank:
        raise ValueError(f"{label} semantic key maps to multiple ranks")
    existing_key = rank_to_key.setdefault(rank, key)
    if existing_key != key:
        raise ValueError(f"{label} rank maps to multiple semantic keys")


def _exact_spatial_int(value: object, label: str) -> int:
    if type(value) is int:
        return value
    if type(value) is float and value.is_integer():
        return int(value)
    raise ValueError(f"spatial destination {label} is not an integer")


def _equal_spatial_pair(left: object, right: object, label: str) -> int:
    left_int = _exact_spatial_int(left, f"{label} minimum")
    right_int = _exact_spatial_int(right, f"{label} maximum")
    if left_int != right_int:
        raise ValueError(f"spatial {label} dimension is not a point")
    return left_int


def _ensure_spatial_dirty_triggers(
    connection: SQLiteConnectionLike,
    backend: EdgeBackend,
    tables: set[str],
) -> None:
    if "graph_spatial_state" not in tables:
        return
    physical_tables = (
        ("edge_rtree_rowid", "edge_source_rtree_rowid")
        if backend == "rtree"
        else ("edge_intervals", "edge_source_intervals")
    )
    for table in physical_tables:
        for operation in ("insert", "update", "delete"):
            name = f"{table}_graph_dirty_{operation}"
            trigger_sql = _expected_graph_trigger_sql(name, table, operation).replace(
                f"CREATE TRIGGER {name}",
                f"CREATE TRIGGER IF NOT EXISTS {name}",
                1,
            )
            connection.execute(trigger_sql)


def _reject_graph_temp_shadows(connection: SQLiteConnectionLike) -> None:
    """Reject TEMP objects that can shadow or intercept authoritative main state."""
    protected = {
        "edges",
        "fblocks",
        "sheets",
        "graph_spatial_state",
        "graph_rank_keys",
        EdgeStore.RTREE_TABLE,
        EdgeStore.SOURCE_RTREE_TABLE,
        EdgeStore.INTERVAL_TABLE,
        EdgeStore.SOURCE_INTERVAL_TABLE,
    }
    protected.update(
        name
        for backend in ("rtree", "interval")
        for name, _table, _operation in _required_graph_triggers(backend)
    )
    for row in connection.execute("SELECT type, name, tbl_name FROM temp.sqlite_master"):
        object_type, raw_name, raw_table = (str(value or "") for value in row)
        name = raw_name.casefold()
        table = raw_table.casefold()
        shadows_rtree = name.startswith(("edge_rtree_", "edge_source_rtree_"))
        targets_rtree = table.startswith(("edge_rtree_", "edge_source_rtree_"))
        if name in protected or table in protected or shadows_rtree or targets_rtree:
            raise EdgeSchemaError(
                f"TEMP {object_type or 'object'} {raw_name!r} shadows graph storage"
            )


def _validate_graph_sidecar(
    connection: SQLiteConnectionLike,
    backend: EdgeBackend,
    tables: set[str],
) -> None:
    """Validate current graph trust state once when opening the sidecar."""
    graph_tables = {
        "edges",
        "fblocks",
        "sheets",
        "graph_spatial_state",
        "graph_rank_keys",
    }
    if "graph_spatial_state" not in tables:
        if tables & graph_tables:
            raise EdgeSchemaError("graph spatial state table is missing")
        # EdgeStore also supports its original standalone range-index mode.
        return
    _validate_physical_backend_schema(connection, backend)
    _validate_rank_key_catalog_schema(connection)

    columns = tuple(connection.execute("PRAGMA main.table_info(graph_spatial_state)"))
    expected_columns = (
        ("singleton", "integer", 0, 1),
        ("dirty", "integer", 1, 0),
        ("dependent_rank_max", "integer", 1, 0),
        ("precedent_rank_max", "integer", 1, 0),
        ("revision", "integer", 1, 0),
        ("mutation_epoch", "integer", 1, 0),
        ("clean_epoch", "integer", 1, 0),
    )
    actual_columns = tuple(
        (str(row[1]), str(row[2]).casefold(), int(row[3]), int(row[5])) for row in columns
    )
    if actual_columns != expected_columns:
        raise EdgeSchemaError("graph spatial state columns do not match the current schema")

    rows = tuple(
        connection.execute(
            """
            SELECT singleton, dirty, dependent_rank_max, precedent_rank_max,
                   revision, mutation_epoch, clean_epoch
            FROM main.graph_spatial_state
            """
        )
    )
    if len(rows) != 1:
        raise EdgeSchemaError("graph spatial state must contain exactly one row")
    try:
        values = tuple(_exact_spatial_int(value, "graph state") for value in rows[0])
    except ValueError as exc:
        raise EdgeSchemaError("graph spatial state contains a non-integer value") from exc
    singleton, dirty, *counters = values
    if singleton != 1 or dirty not in (0, 1) or any(counter < 0 for counter in counters):
        raise EdgeSchemaError("graph spatial state contains an invalid value")
    if values[-1] > values[-2]:
        raise EdgeSchemaError("graph spatial clean epoch is ahead of its mutation epoch")
    try:
        canonical_records = canonical_ranked_edges(connection, require_persisted_ranks=True)
    except GraphProjectionError as exc:
        raise EdgeSchemaError(
            "relational graph ranks do not match canonical public-hop semantics"
        ) from exc
    if values[2] != max(
        (record.dependent_rank for record in canonical_records), default=0
    ) or values[3] != max((record.precedent_rank for record in canonical_records), default=0):
        raise EdgeSchemaError("graph spatial rank maxima do not match canonical public-hop ranks")
    _validate_full_spatial_mirrors(connection, backend)
    expected_rank_keys: dict[tuple[str, int], str] = {}
    for record in canonical_records:
        expected_rank_keys[("dependents", record.dependent_rank)] = canonical_rank_key_text(
            record.dependent_key
        )
        expected_rank_keys[("precedents", record.precedent_rank)] = canonical_rank_key_text(
            record.precedent_key
        )
    actual_rank_keys = {
        (str(row[0]), _exact_spatial_int(row[1], "rank key rank")): str(row[2])
        for row in connection.execute(
            "SELECT direction, rank, key_text FROM main.graph_rank_keys ORDER BY direction, rank"
        )
    }
    if actual_rank_keys != expected_rank_keys:
        raise EdgeSchemaError("graph rank-key catalog does not match canonical public hops")

    required = _required_graph_triggers(backend)
    trigger_rows = {
        str(row[0]): (str(row[1]), str(row[2] or ""))
        for row in connection.execute(
            "SELECT name, tbl_name, sql FROM main.sqlite_master WHERE type = 'trigger'"
        )
    }
    expected_names = {name for name, _table, _operation in required}
    protected_tables = {
        "edges",
        "fblocks",
        "sheets",
        "graph_spatial_state",
        *(table for _name, table, _operation in required),
    }
    for name, (table, sql) in trigger_rows.items():
        if name not in expected_names and (
            table in protected_tables or "graph_spatial_state" in sql.casefold()
        ):
            raise EdgeSchemaError(f"unexpected graph state trigger {name} is present")
    for name, table, operation in required:
        trigger = trigger_rows.get(name)
        if trigger is None or trigger[0] != table:
            raise EdgeSchemaError(f"required graph dirty trigger {name} is missing")
        if _canonicalize_trigger_sql(trigger[1]) != _canonicalize_trigger_sql(
            _expected_graph_trigger_sql(name, table, operation)
        ):
            raise EdgeSchemaError(f"required graph dirty trigger {name} is malformed")


def _validate_physical_backend_schema(
    connection: SQLiteConnectionLike,
    backend: EdgeBackend,
) -> None:
    """Validate exact main-table columns and RTree virtual-table identity."""
    try:
        if backend == "rtree":
            columns = (
                "edge_id",
                "sheet_min",
                "sheet_max",
                "row_min",
                "row_max",
                "col_min",
                "col_max",
                "rank_min",
                "rank_max",
            )
            expected_info = tuple((name, "int", 0, 0) for name in columns)
            for table in (EdgeStore.RTREE_TABLE, EdgeStore.SOURCE_RTREE_TABLE):
                _validate_one_physical_table(
                    connection,
                    table,
                    expected_info,
                    f"CREATE VIRTUAL TABLE {table} USING rtree_i32({','.join(columns)})",
                )
                _validate_rtree_shadow_storage(connection, table)
            return

        expected_info = (
            ("edge_id", "integer", 0, 1),
            ("sheet_id", "integer", 1, 0),
            ("row_min", "integer", 1, 0),
            ("row_max", "integer", 1, 0),
            ("col_min", "integer", 1, 0),
            ("col_max", "integer", 1, 0),
            ("rank", "integer", 1, 0),
        )
        for table in (EdgeStore.INTERVAL_TABLE, EdgeStore.SOURCE_INTERVAL_TABLE):
            _validate_one_physical_table(
                connection,
                table,
                expected_info,
                f"""
                CREATE TABLE {table} (
                    edge_id INTEGER PRIMARY KEY,
                    sheet_id INTEGER NOT NULL,
                    row_min INTEGER NOT NULL,
                    row_max INTEGER NOT NULL,
                    col_min INTEGER NOT NULL,
                    col_max INTEGER NOT NULL,
                    rank INTEGER NOT NULL
                )
                """,
            )
    except sqlite3.DatabaseError as exc:
        raise EdgeSchemaError("graph spatial physical schema could not be inspected") from exc


def _validate_rank_key_catalog_schema(connection: SQLiteConnectionLike) -> None:
    """Validate the exact bounded semantic-rank catalog identity."""
    expected_info = (
        ("direction", "text", 1, 1),
        ("rank", "integer", 1, 2),
        ("key_text", "text", 1, 0),
    )
    _validate_one_physical_table(
        connection,
        "graph_rank_keys",
        expected_info,
        """
        CREATE TABLE graph_rank_keys (
            direction TEXT NOT NULL CHECK (direction IN ('dependents', 'precedents')),
            rank INTEGER NOT NULL CHECK (rank > 0),
            key_text TEXT NOT NULL,
            PRIMARY KEY (direction, rank)
        ) WITHOUT ROWID
        """,
    )


def _validate_one_physical_table(
    connection: SQLiteConnectionLike,
    table: str,
    expected_info: tuple[tuple[str, str, int, int], ...],
    expected_sql: str,
) -> None:
    row = connection.execute(
        "SELECT type, sql FROM main.sqlite_master WHERE name = ?",
        (table,),
    ).fetchone()
    if row is None or str(row[0]).casefold() != "table" or row[1] is None:
        raise EdgeSchemaError(f"graph spatial table {table} is missing or has the wrong identity")
    actual_info = tuple(
        (str(item[1]), str(item[2]).casefold(), int(item[3]), int(item[5]))
        for item in connection.execute(f"PRAGMA main.table_info({table})")
    )
    if actual_info != expected_info:
        raise EdgeSchemaError(f"graph spatial table {table} has malformed columns")
    if _canonicalize_schema_sql(str(row[1])) != _canonicalize_schema_sql(expected_sql):
        raise EdgeSchemaError(f"graph spatial table {table} has the wrong physical identity")


def _validate_rtree_shadow_storage(connection: SQLiteConnectionLike, table: str) -> None:
    messages = tuple(
        str(row[0]) for row in connection.execute(f"PRAGMA main.integrity_check('{table}')")
    )
    if messages != ("ok",):
        raise EdgeSchemaError(f"graph spatial table {table} has corrupt shadow storage")


def _validate_full_spatial_mirrors(
    connection: SQLiteConnectionLike,
    backend: EdgeBackend,
) -> None:
    """Compare both complete persisted mirrors with their relational projection."""
    source_expected = f"""
        SELECT e.id,
               e.src_sheet_id,
               {"e.src_sheet_id," if backend == "rtree" else ""}
               CASE
                   WHEN e.src_kind = 'fblock' THEN fb.row_min
                   WHEN e.src_kind = 'cell' THEN CAST(e.src_id / 65536 AS INTEGER)
               END,
               CASE
                   WHEN e.src_kind = 'fblock' THEN fb.row_max
                   WHEN e.src_kind = 'cell' THEN CAST(e.src_id / 65536 AS INTEGER)
               END,
               CASE
                   WHEN e.src_kind = 'fblock' THEN fb.col_min
                   WHEN e.src_kind = 'cell' THEN (e.src_id & 65535)
               END,
               CASE
                   WHEN e.src_kind = 'fblock' THEN fb.col_max
                   WHEN e.src_kind = 'cell' THEN (e.src_id & 65535)
               END,
               e.precedent_rank
               {", e.precedent_rank" if backend == "rtree" else ""}
        FROM main.edges AS e
        LEFT JOIN main.fblocks AS fb
               ON e.src_kind = 'fblock'
              AND fb.id = e.src_id
              AND fb.sheet_id = e.src_sheet_id
    """
    destination_expected = f"""
        SELECT e.id, e.dst_sheet_id,
               {"e.dst_sheet_id," if backend == "rtree" else ""}
               e.dst_row_min, e.dst_row_max, e.dst_col_min, e.dst_col_max,
               e.dependent_rank
               {", e.dependent_rank" if backend == "rtree" else ""}
        FROM main.edges AS e
        WHERE e.dst_sheet_id IS NOT NULL
          AND e.dst_row_min IS NOT NULL AND e.dst_row_max IS NOT NULL
          AND e.dst_col_min IS NOT NULL AND e.dst_col_max IS NOT NULL
    """
    if backend == "rtree":
        source_actual = """
            SELECT edge_id, sheet_min, sheet_max, row_min, row_max,
                   col_min, col_max, rank_min, rank_max
            FROM main.edge_source_rtree
        """
        destination_actual = """
            SELECT edge_id, sheet_min, sheet_max, row_min, row_max,
                   col_min, col_max, rank_min, rank_max
            FROM main.edge_rtree
        """
    else:
        source_actual = """
            SELECT edge_id, sheet_id, row_min, row_max, col_min, col_max, rank
            FROM main.edge_source_intervals
        """
        destination_actual = """
            SELECT edge_id, sheet_id, row_min, row_max, col_min, col_max, rank
            FROM main.edge_intervals
        """
    try:
        partial_destination = connection.execute(
            """
            SELECT 1 FROM main.edges
            WHERE NOT (
                (dst_sheet_id IS NULL AND dst_row_min IS NULL AND dst_row_max IS NULL
                 AND dst_col_min IS NULL AND dst_col_max IS NULL)
                OR
                (dst_sheet_id IS NOT NULL AND dst_row_min IS NOT NULL
                 AND dst_row_max IS NOT NULL AND dst_col_min IS NOT NULL
                 AND dst_col_max IS NOT NULL)
            )
            LIMIT 1
            """
        ).fetchone()
        if partial_destination is not None:
            raise EdgeSchemaError("relational dependency has a partial destination rectangle")
        for label, expected, actual in (
            ("source", source_expected, source_actual),
            ("destination", destination_expected, destination_actual),
        ):
            if _sql_sets_differ(connection, expected, actual):
                raise EdgeSchemaError(
                    f"persisted {label} spatial mirror does not match relational dependencies"
                )
    except sqlite3.DatabaseError as exc:
        raise EdgeSchemaError("persisted spatial mirrors could not be validated") from exc


def _sql_sets_differ(connection: SQLiteConnectionLike, left: str, right: str) -> bool:
    for first, second in ((left, right), (right, left)):
        row = connection.execute(f"SELECT 1 FROM ({first} EXCEPT {second}) LIMIT 1").fetchone()
        if row is not None:
            return True
    return False


def _canonicalize_trigger_sql(sql: str) -> str:
    """Return exact trigger text modulo case and insignificant whitespace."""
    return _canonicalize_schema_sql(sql)


def _canonicalize_schema_sql(sql: str) -> str:
    return "".join(sql.casefold().split())


def _expected_graph_trigger_sql(name: str, table: str, operation: str) -> str:
    invalidation = _rank_catalog_invalidation_sql(table, operation)
    return f"""
        CREATE TRIGGER {name}
        AFTER {operation.upper()} ON {table}
        BEGIN
            {invalidation}
            UPDATE graph_spatial_state
            SET dirty = 1, mutation_epoch = mutation_epoch + 1
            WHERE singleton = 1;
        END
    """


def _rank_catalog_invalidation_sql(table: str, operation: str) -> str:
    if table == "edges":
        if operation == "insert":
            return """
                DELETE FROM graph_rank_keys
                WHERE (direction = 'dependents' AND rank = NEW.dependent_rank)
                   OR (direction = 'precedents' AND rank = NEW.precedent_rank);
            """
        if operation == "update":
            return """
                DELETE FROM graph_rank_keys
                WHERE (direction = 'dependents'
                       AND rank IN (OLD.dependent_rank, NEW.dependent_rank))
                   OR (direction = 'precedents'
                       AND rank IN (OLD.precedent_rank, NEW.precedent_rank));
            """
        return """
            DELETE FROM graph_rank_keys
            WHERE (direction = 'dependents' AND rank = OLD.dependent_rank)
               OR (direction = 'precedents' AND rank = OLD.precedent_rank);
        """
    if table in {"fblocks", "sheets"}:
        return "DELETE FROM graph_rank_keys;"
    mirror_direction = (
        "precedents"
        if table in {"edge_source_intervals", "edge_source_rtree_rowid"}
        else "dependents"
        if table in {"edge_intervals", "edge_rtree_rowid"}
        else None
    )
    if mirror_direction is not None:
        if table.endswith("_intervals"):
            if operation == "insert":
                rank_predicate = "rank = NEW.rank"
            elif operation == "update":
                rank_predicate = "rank IN (OLD.rank, NEW.rank)"
            else:
                rank_predicate = "rank = OLD.rank"
            return (
                "DELETE FROM graph_rank_keys "
                f"WHERE direction = '{mirror_direction}' AND {rank_predicate};"
            )
        # RTree virtual tables cannot own triggers; their rowid shadow does
        # not expose rank, so invalidate the bounded direction catalog.
        return f"DELETE FROM graph_rank_keys WHERE direction = '{mirror_direction}';"
    return ""


def _required_graph_triggers(
    backend: EdgeBackend,
) -> tuple[tuple[str, str, str], ...]:
    base_tables = ("edges", "fblocks", "sheets", "graph_rank_keys")
    mirror_tables = (
        ("edge_rtree_rowid", "edge_source_rtree_rowid")
        if backend == "rtree"
        else ("edge_intervals", "edge_source_intervals")
    )
    return (
        *(
            (f"{table}_graph_spatial_dirty_{operation}", table, operation)
            for table in base_tables
            for operation in ("insert", "update", "delete")
        ),
        *(
            (f"{table}_graph_dirty_{operation}", table, operation)
            for table in mirror_tables
            for operation in ("insert", "update", "delete")
        ),
    )


__all__ = [
    "EdgeBackend",
    "EdgeDirection",
    "EdgeSchemaError",
    "EdgeStore",
    "GraphProjectionError",
    "RankedEdge",
    "SQLiteConnectionLike",
    "canonical_rank_key_text",
    "canonical_ranked_edges",
]
