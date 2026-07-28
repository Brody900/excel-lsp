"""Deterministic, bounded dependency queries over persisted block-level edges."""

from __future__ import annotations

import sqlite3
from collections import deque
from collections.abc import Callable, Generator, Iterable, Sequence
from contextlib import contextmanager
from dataclasses import dataclass, field
from functools import wraps
from typing import Any, NoReturn, ParamSpec, TypeAlias, TypeVar, cast

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.exception_evidence import (
    normalize_exception_graph,
    prepare_chained_failure_with_primary_evidence,
)
from excel_lsp.core.graph.models import (
    GraphArea,
    GraphDirection,
    GraphHop,
    GraphTarget,
    PathResult,
    PathStep,
    TraceNode,
    TraceResult,
)
from excel_lsp.core.index.edges import (
    EdgeSchemaError,
    EdgeStore,
    SQLiteConnectionLike,
    canonical_rank_key_text,
)
from excel_lsp.core.models import Rect
from excel_lsp.core.symbols import cell_symbol_id, formula_block_symbol_id

_MAX_TRACE_DEPTH = 8
_MAX_TRACE_NODES = 10_000
_MAX_PATHS = 50
_MAX_PATH_DEPTH = 12
_MAX_PATH_NODES = 10_000
_CELL_ID_FACTOR = 1 << 16
_ROW_MAX = 1_048_576
_COL_MAX = 16_384
_SQL_ID_CHUNK = 400

_NodeKey: TypeAlias = tuple[str, int, int, int, int, int, str]
_QueryParameters = ParamSpec("_QueryParameters")
_QueryResult = TypeVar("_QueryResult")


def _cleanup_failure(errors: tuple[BaseException, ...]) -> BaseException:
    failure: BaseException = (
        errors[0]
        if len(errors) == 1
        else BaseExceptionGroup("dependency graph snapshot cleanup failures", errors)
    )
    normalized = normalize_exception_graph(failure)
    if normalized is None:
        raise AssertionError("cleanup evidence unexpectedly normalized to empty")
    return normalized


def _raise_cleanup_failure(errors: tuple[BaseException, ...]) -> NoReturn:
    primary_error = _cleanup_failure((errors[0],))
    if len(errors) > 1:
        primary_error.add_note("Additional dependency graph snapshot cleanup failures are chained.")
        cleanup_failure = prepare_chained_failure_with_primary_evidence(
            _cleanup_failure(errors[1:]),
            primary_error,
            message=(
                "Dependency graph primary cleanup causal evidence and additional cleanup failures"
            ),
        )
        if cleanup_failure is not None:
            raise primary_error from cleanup_failure
    raise primary_error


def _shape_query_database_errors(
    method: Callable[_QueryParameters, _QueryResult],
) -> Callable[_QueryParameters, _QueryResult]:
    """Translate live SQLite failures at one complete public query boundary."""

    @wraps(method)
    def wrapped(*args: _QueryParameters.args, **kwargs: _QueryParameters.kwargs) -> _QueryResult:
        try:
            return method(*args, **kwargs)
        except sqlite3.ProgrammingError:
            raise
        except sqlite3.DatabaseError as exc:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Dependency graph storage became corrupt while executing a query.",
            ) from exc

    return wrapped


@dataclass(slots=True)
class _MutableTraceNode:
    target: GraphTarget
    via: str | None
    children: list[_MutableTraceNode] = field(default_factory=list["_MutableTraceNode"])
    child_count: int = 0


class DependencyGraph:
    """Query one index connection without constructing a cell-level graph."""

    def __init__(
        self,
        connection: SQLiteConnectionLike,
        edge_store: EdgeStore | None = None,
    ) -> None:
        self._connection = connection
        self._snapshot_poisoned = False
        try:
            self._edges = edge_store or EdgeStore(connection)
        except sqlite3.ProgrammingError:
            raise
        except (EdgeSchemaError, sqlite3.DatabaseError, ValueError) as exc:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Dependency graph storage schema is missing or corrupt.",
            ) from exc

    @_shape_query_database_errors
    def direct_precedents(self, query: GraphArea | GraphTarget) -> tuple[GraphHop, ...]:
        """Return destinations read by source blocks intersecting ``query``."""
        with self._consistent_read_snapshot():
            self._require_clean_spatial()
            area = _require_area(query)
            self._validate_source_query(area)
            return _deduplicate_hops(
                self._precedent_hop(row) for row in self._iter_source_overlap_rows(area)
            )

    @_shape_query_database_errors
    def direct_dependents(self, query: GraphArea | GraphTarget) -> tuple[GraphHop, ...]:
        """Return source blocks whose destination rectangles intersect ``query``."""
        with self._consistent_read_snapshot():
            self._require_clean_spatial()
            area = _require_area(query)
            return _deduplicate_hops(self._iter_dependent_hops(area))

    @_shape_query_database_errors
    def trace_precedents(
        self,
        query: GraphArea | GraphTarget,
        *,
        depth: int = 2,
        max_nodes: int = 200,
    ) -> TraceResult:
        """Trace upstream references as a bounded breadth-first tree."""
        with self._consistent_read_snapshot():
            return self._trace(query, "precedents", depth=depth, max_nodes=max_nodes)

    @_shape_query_database_errors
    def trace_dependents(
        self,
        query: GraphArea | GraphTarget,
        *,
        depth: int = 2,
        max_nodes: int = 200,
    ) -> TraceResult:
        """Trace downstream formula blocks as a bounded breadth-first tree."""
        with self._consistent_read_snapshot():
            return self._trace(query, "dependents", depth=depth, max_nodes=max_nodes)

    @_shape_query_database_errors
    def trace_path(
        self,
        source: GraphArea | GraphTarget,
        destination: GraphArea | GraphTarget,
        *,
        max_paths: int = 3,
        max_depth: int = 12,
    ) -> PathResult:
        """Return bounded shortest block-level paths in dependent direction."""
        with self._consistent_read_snapshot():
            return self._trace_path(
                source,
                destination,
                max_paths=max_paths,
                max_depth=max_depth,
            )

    def _trace_path(
        self,
        source: GraphArea | GraphTarget,
        destination: GraphArea | GraphTarget,
        *,
        max_paths: int,
        max_depth: int,
    ) -> PathResult:
        self._require_clean_spatial()
        _validate_bounded_int("max_paths", max_paths, minimum=1, maximum=_MAX_PATHS)
        _validate_bounded_int("max_depth", max_depth, minimum=0, maximum=_MAX_PATH_DEPTH)
        source_target = _root_target(source)
        source_area = _require_area(source_target)
        destination_area = _require_area(destination)
        source_key = _target_key(source_target)
        if _areas_intersect(source_area, destination_area):
            return PathResult(
                True,
                ((PathStep(source_target.label, None),),),
            )

        queue: deque[_NodeKey] = deque((source_key,))
        targets = {source_key: source_target}
        distances = {source_key: 0}
        parents: dict[_NodeKey, list[tuple[_NodeKey, str]]] = {source_key: []}
        destination_keys: set[_NodeKey] = set()
        destination_depth: int | None = None
        truncated = False

        while queue:
            node_key = queue.popleft()
            node_depth = distances[node_key]
            if node_depth >= max_depth:
                continue
            if destination_depth is not None and node_depth >= destination_depth:
                continue
            node_target = targets[node_key]
            remaining_nodes = _MAX_PATH_NODES - len(targets)
            hops, scan_truncated = self._bounded_hops(
                node_target,
                "dependents",
                limit=remaining_nodes + max_paths + 1,
            )
            truncated = truncated or scan_truncated
            if len(hops) > remaining_nodes + max_paths:
                truncated = True
            for hop in hops[: remaining_nodes + max_paths]:
                child = hop.target
                child_area = child.area
                if child_area is None:
                    continue
                child_key = _target_key(child)
                child_depth = node_depth + 1
                known_depth = distances.get(child_key)
                if known_depth is None:
                    if len(targets) >= _MAX_PATH_NODES:
                        truncated = True
                        continue
                    distances[child_key] = child_depth
                    targets[child_key] = child
                    parents[child_key] = [(node_key, hop.via)]
                    queue.append(child_key)
                elif known_depth == child_depth:
                    parent = (node_key, hop.via)
                    if parent not in parents[child_key]:
                        if len(parents[child_key]) < max_paths:
                            parents[child_key].append(parent)
                        else:
                            truncated = True
                else:
                    continue
                if _areas_intersect(child_area, destination_area):
                    if destination_depth is None:
                        destination_depth = child_depth
                    if child_depth == destination_depth:
                        destination_keys.add(child_key)

        paths: list[tuple[PathStep, ...]] = []
        ordered_destinations = sorted(destination_keys)
        for destination_index, destination_key in enumerate(ordered_destinations):
            remaining_paths = max_paths - len(paths)
            reconstructed, combinations_truncated = _reconstruct_paths(
                destination_key, source_key, parents, remaining_paths
            )
            truncated = truncated or combinations_truncated
            for key_path in reconstructed:
                steps: list[PathStep] = []
                for key, incoming in key_path:
                    steps.append(PathStep(targets[key].label, incoming))
                candidate = tuple(steps)
                if candidate not in paths:
                    paths.append(candidate)
                if len(paths) >= max_paths:
                    truncated = truncated or destination_index + 1 < len(ordered_destinations)
                    return PathResult(True, tuple(paths), truncated)
        return PathResult(bool(paths), tuple(paths), truncated)

    def _trace(
        self,
        query: GraphArea | GraphTarget,
        direction: GraphDirection,
        *,
        depth: int,
        max_nodes: int,
    ) -> TraceResult:
        self._require_clean_spatial()
        _validate_bounded_int("depth", depth, minimum=0, maximum=_MAX_TRACE_DEPTH)
        _validate_bounded_int("max_nodes", max_nodes, minimum=1, maximum=_MAX_TRACE_NODES)
        root_target = _root_target(query)
        root = _MutableTraceNode(root_target, None)
        queue: deque[tuple[_MutableTraceNode, int]] = deque(((root, 0),))
        expanded = {_target_key(root_target)}
        node_count = 1
        edge_count = 0
        truncated = False

        while queue:
            node, node_depth = queue.popleft()
            if node_depth >= depth:
                continue
            remaining = max_nodes - node_count
            hops, scan_truncated = self._bounded_hops(node.target, direction, limit=remaining + 1)
            truncated = truncated or scan_truncated
            node.child_count = len(hops)
            if len(hops) > remaining:
                truncated = True
            for hop in hops[:remaining]:
                child = _MutableTraceNode(hop.target, hop.via)
                node.children.append(child)
                node_count += 1
                edge_count += 1
                child_key = _target_key(hop.target)
                if hop.target.area is not None and child_key not in expanded:
                    expanded.add(child_key)
                    queue.append((child, node_depth + 1))
            if truncated:
                break

        return TraceResult(
            direction,
            _freeze_trace(root),
            node_count,
            edge_count,
            truncated,
        )

    def _iter_source_overlap_rows(
        self, area: GraphArea, *, limit: int | None = None
    ) -> Iterable[dict[str, Any]]:
        sql = f"""
            {_EDGE_SELECT}
            WHERE e.src_sheet_id = ? AND (
                (e.src_kind = 'fblock'
                 AND fb.row_min <= ? AND fb.row_max >= ?
                 AND fb.col_min <= ? AND fb.col_max >= ?)
                OR
                (e.src_kind = 'cell'
                 AND CAST(e.src_id / {_CELL_ID_FACTOR} AS INTEGER) BETWEEN ? AND ?
                 AND (e.src_id & {_CELL_ID_FACTOR - 1}) BETWEEN ? AND ?)
            )
            ORDER BY e.id
            """
        parameters: tuple[object, ...] = (
            area.sheet_id,
            area.rect.row_max,
            area.rect.row_min,
            area.rect.col_max,
            area.rect.col_min,
            area.rect.row_min,
            area.rect.row_max,
            area.rect.col_min,
            area.rect.col_max,
        )
        if limit is not None:
            sql += " LIMIT ?"
            parameters = (*parameters, limit)
        return self._iter_rows(sql, parameters)

    def _edge_rows(self, edge_ids: Sequence[int]) -> tuple[dict[str, Any], ...]:
        rows: list[dict[str, Any]] = []
        for start in range(0, len(edge_ids), _SQL_ID_CHUNK):
            chunk = edge_ids[start : start + _SQL_ID_CHUNK]
            placeholders = ",".join("?" for _edge_id in chunk)
            rows.extend(
                self._rows(
                    f"{_EDGE_SELECT} WHERE e.id IN ({placeholders})",
                    tuple(chunk),
                )
            )
        return tuple(rows)

    def _precedent_hop(self, row: dict[str, Any]) -> GraphHop:
        _source_target(row)
        destination = _destination_target(row)
        self._validate_spatial_mirror(row)
        return GraphHop(destination, str(row["via"]))

    def _dependent_hop(self, row: dict[str, Any]) -> GraphHop:
        source = _source_target(row)
        _destination_target(row)
        self._validate_spatial_mirror(row)
        return GraphHop(source, str(row["via"]))

    def _validate_spatial_mirror(self, row: dict[str, Any]) -> None:
        edge_id = _require_integral(row["edge_id"], 0, "edge id")
        try:
            destination_spatial = self._edges.ranked_mirror(edge_id, "dependents")
            source_spatial = self._edges.ranked_mirror(edge_id, "precedents")
        except sqlite3.ProgrammingError:
            raise
        except (sqlite3.DatabaseError, ValueError) as exc:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                f"Dependency edge {edge_id} has an invalid spatial mirror.",
            ) from exc
        source = _source_target(row)
        source_rank = _require_integral(row["precedent_rank"], edge_id, "precedent rank")
        expected_source = (
            source.area.sheet_id if source.area is not None else None,
            source.area.rect if source.area is not None else None,
            source_rank,
        )
        if source_spatial != expected_source:
            _corrupt(edge_id, "relational source does not match its ranked spatial mirror")

        destination_values = (
            row["dst_sheet_id"],
            row["dst_row_min"],
            row["dst_row_max"],
            row["dst_col_min"],
            row["dst_col_max"],
        )
        if all(value is None for value in destination_values):
            if destination_spatial is not None:
                _corrupt(edge_id, "opaque destination unexpectedly has a spatial mirror")
            return
        expected = (
            _require_integral(row["dst_sheet_id"], edge_id, "destination sheet id"),
            _rect_from_values(
                edge_id,
                row["dst_row_min"],
                row["dst_row_max"],
                row["dst_col_min"],
                row["dst_col_max"],
                label="destination",
            ),
            _require_integral(row["dependent_rank"], edge_id, "dependent rank"),
        )
        if destination_spatial != expected:
            _corrupt(
                edge_id,
                "relational destination does not match its ranked spatial mirror",
            )

    def _rows(self, sql: str, parameters: tuple[object, ...]) -> tuple[dict[str, Any], ...]:
        return tuple(self._iter_rows(sql, parameters))

    def _iter_rows(self, sql: str, parameters: tuple[object, ...]) -> Iterable[dict[str, Any]]:
        cursor = self._connection.execute(sql, parameters)
        names = tuple(item[0] for item in cursor.description or ())
        for row in cursor:
            yield dict(zip(names, tuple(row), strict=True))

    def _iter_dependent_hops(self, area: GraphArea) -> Iterable[GraphHop]:
        self._validate_destination_query(area)
        after_edge_id = 0
        while page := self._edges.query_range_page(
            area.sheet_id,
            area.rect,
            after_edge_id=after_edge_id,
            limit=EdgeStore.MAX_PAGE_SIZE,
        ):
            rows = self._edge_rows(page)
            self._require_relational_rows(page, rows)
            hops = tuple(self._dependent_hop(row) for row in rows)
            yield from sorted(hops, key=_hop_sort_key)
            after_edge_id = page[-1]

    def _bounded_hops(
        self,
        query: GraphArea | GraphTarget,
        direction: GraphDirection,
        *,
        limit: int,
    ) -> tuple[tuple[GraphHop, ...], bool]:
        """Read the exact semantic prefix through the ranked spatial mirrors."""
        area = _require_area(query)
        self._require_clean_spatial()
        hops: list[GraphHop] = []
        after_rank = 0
        while len(hops) < limit:
            try:
                rank = self._edges.first_matching_rank(
                    direction,
                    area.sheet_id,
                    area.rect,
                    after_rank=after_rank,
                )
            except sqlite3.ProgrammingError:
                raise
            except (RuntimeError, sqlite3.DatabaseError, ValueError) as exc:
                raise ExcelLSPError(
                    ErrorCode.CORRUPT,
                    "Dependency graph ranked spatial mirrors are invalid.",
                ) from exc
            if rank is None:
                break
            try:
                edge_id = self._edges.edge_id_at_rank(direction, area.sheet_id, area.rect, rank)
            except sqlite3.ProgrammingError:
                raise
            except (RuntimeError, sqlite3.DatabaseError, ValueError) as exc:
                raise ExcelLSPError(
                    ErrorCode.CORRUPT,
                    "Dependency graph ranked spatial mirrors are invalid.",
                ) from exc
            if edge_id is None:
                raise ExcelLSPError(
                    ErrorCode.CORRUPT,
                    f"Dependency graph rank {rank} has no representative edge.",
                )
            rows = self._edge_rows((edge_id,))
            self._require_relational_rows((edge_id,), rows)
            if len(rows) != 1:
                raise ExcelLSPError(
                    ErrorCode.CORRUPT,
                    f"Dependency graph rank {rank} has an ambiguous representative edge.",
                )
            row = rows[0]
            expected_rank_column = (
                "precedent_rank" if direction == "precedents" else "dependent_rank"
            )
            if _require_integral(row[expected_rank_column], edge_id, expected_rank_column) != rank:
                _corrupt(edge_id, f"{expected_rank_column} does not match its spatial mirror")
            hop = (
                self._precedent_hop(row) if direction == "precedents" else self._dependent_hop(row)
            )
            try:
                expected_key = self._edges.rank_key_text(direction, rank)
            except sqlite3.ProgrammingError:
                raise
            except (RuntimeError, sqlite3.DatabaseError, ValueError) as exc:
                raise ExcelLSPError(
                    ErrorCode.CORRUPT,
                    f"Dependency graph rank {rank} has no trustworthy semantic identity.",
                ) from exc
            if canonical_rank_key_text(_hop_sort_key(hop)) != expected_key:
                _corrupt(edge_id, f"public hop does not match the catalog for rank {rank}")
            hops.append(hop)
            after_rank = rank
        return tuple(hops), False

    def _require_clean_spatial(self) -> None:
        try:
            self._edges.require_clean()
        except sqlite3.ProgrammingError:
            raise
        except (RuntimeError, sqlite3.DatabaseError, ValueError) as exc:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Dependency graph ranked spatial mirrors are stale or corrupt.",
            ) from exc

    @contextmanager
    def _consistent_read_snapshot(self) -> Generator[None, None, None]:
        """Keep one public graph operation on a single SQLite read snapshot."""
        if self._snapshot_poisoned:
            self._raise_poisoned_connection()
        owns_snapshot = not self._connection.in_transaction
        if not owns_snapshot:
            yield
            return

        try:
            # Connection subclasses can raise either before or after the
            # native BEGIN takes effect.  Acquisition therefore belongs to
            # the same cleanup scope as the query body.
            self._connection.execute("BEGIN")
            yield
        except BaseException as primary_error:
            cleanup_errors = self._release_owned_read_snapshot()
            if cleanup_errors:
                primary_error.add_note(
                    "Dependency graph read-snapshot cleanup also failed; "
                    "the query error remains primary."
                )
                cleanup_failure = _cleanup_failure(cleanup_errors)
                cleanup_failure = prepare_chained_failure_with_primary_evidence(
                    cleanup_failure,
                    primary_error,
                    message=("Dependency graph body causal evidence and snapshot cleanup failure"),
                )
                if cleanup_failure is not None:
                    raise primary_error from cleanup_failure
            raise
        else:
            cleanup_errors = self._release_owned_read_snapshot()
            if cleanup_errors:
                _raise_cleanup_failure(cleanup_errors)

    def _rollback_owned_read_snapshot(self) -> None:
        self._connection.rollback()

    def _release_owned_read_snapshot(self) -> tuple[BaseException, ...]:
        """Release an owned snapshot or poison its connection conclusively."""
        cleanup_errors: list[BaseException] = []
        try:
            self._rollback_owned_read_snapshot()
        except BaseException as cleanup_error:
            cleanup_errors.append(cleanup_error)

        state_known = True
        try:
            still_active = self._connection.in_transaction
        except BaseException as state_error:
            cleanup_errors.append(state_error)
            state_known = False
            still_active = True

        if state_known and still_active:
            try:
                self._connection.execute("ROLLBACK")
            except BaseException as fallback_error:
                cleanup_errors.append(fallback_error)
            try:
                still_active = self._connection.in_transaction
            except BaseException as state_error:
                cleanup_errors.append(state_error)
                state_known = False
                still_active = True

        if not state_known or still_active:
            self._snapshot_poisoned = True
            self._conclusively_close_poisoned_connection(cleanup_errors)
        return tuple(cleanup_errors)

    def _close_poisoned_connection(self) -> None:
        self._connection.close()

    def _raise_poisoned_connection(self) -> NoReturn:
        """Retry physical closure before rejecting reuse of a poisoned graph."""
        cleanup_errors: list[BaseException] = []
        self._conclusively_close_poisoned_connection(cleanup_errors)
        poisoned_error = sqlite3.ProgrammingError("dependency graph connection is poisoned")
        if cleanup_errors:
            raise poisoned_error from _cleanup_failure(tuple(cleanup_errors))
        raise poisoned_error

    def _conclusively_close_poisoned_connection(self, cleanup_errors: list[BaseException]) -> None:
        """Close a poisoned connection, retrying independently after a failed hook."""
        close_attempts = (
            self._close_poisoned_connection,
            self._connection.close,
            self._connection.close,
        )
        for close_connection in close_attempts:
            try:
                close_connection()
            except BaseException as close_error:
                cleanup_errors.append(close_error)
            if self._connection_is_physically_closed(cleanup_errors):
                return

        # A real sqlite3.Connection may be a testable/instrumented subclass
        # whose close override failed before reaching the native descriptor.
        # The base descriptor is the final supported emergency release.
        try:
            connection = cast(object, self._connection)
            if isinstance(connection, sqlite3.Connection):
                sqlite3.Connection.close(connection)
            else:
                self._connection.close()
        except BaseException as close_error:
            cleanup_errors.append(close_error)
        self._connection_is_physically_closed(cleanup_errors)

    def _connection_is_physically_closed(self, cleanup_errors: list[BaseException]) -> bool:
        """Probe native descriptor state without trusting subclass overrides."""
        try:
            connection = cast(object, self._connection)
            if isinstance(connection, sqlite3.Connection):
                descriptor = cast(Any, sqlite3.Connection.in_transaction)
                _ = descriptor.__get__(connection, sqlite3.Connection)
            else:
                _ = self._connection.in_transaction
        except sqlite3.ProgrammingError as state_error:
            if "closed" in str(state_error).casefold():
                return True
            cleanup_errors.append(state_error)
        except BaseException as state_error:
            cleanup_errors.append(state_error)
        return False

    def _validate_source_query(self, area: GraphArea) -> None:
        row = self._connection.execute(
            f"""
            SELECT e.id, e.src_kind
            FROM main.edges AS e
            LEFT JOIN main.fblocks AS fb
                   ON e.src_kind = 'fblock' AND fb.id = e.src_id
            WHERE e.src_sheet_id = ? AND (
                e.src_kind NOT IN ('fblock', 'cell')
                OR (e.src_kind = 'fblock'
                    AND (fb.id IS NULL OR fb.sheet_id != e.src_sheet_id))
                OR (e.src_kind = 'cell' AND (
                    typeof(e.src_id) != 'integer'
                    OR CAST(e.src_id / {_CELL_ID_FACTOR} AS INTEGER)
                       NOT BETWEEN 1 AND {_ROW_MAX}
                    OR (e.src_id & {_CELL_ID_FACTOR - 1}) NOT BETWEEN 1 AND {_COL_MAX}
                ))
            )
            ORDER BY e.id
            LIMIT 1
            """,
            (area.sheet_id,),
        ).fetchone()
        if row is not None:
            edge_id = _require_integral(row[0], 0, "edge id")
            _corrupt(edge_id, f"unsupported or orphaned source kind {row[1]!r}")

    def _validate_destination_query(self, area: GraphArea) -> None:
        """Reject matching relational destinations with absent or wrong mirrors."""
        if self._edges.backend == "rtree":
            spatial_join = "LEFT JOIN main.edge_rtree AS spatial ON spatial.edge_id = e.id"
            mismatch = """
                spatial.edge_id IS NULL
                OR spatial.sheet_min != e.dst_sheet_id
                OR spatial.sheet_max != e.dst_sheet_id
                OR spatial.row_min != e.dst_row_min
                OR spatial.row_max != e.dst_row_max
                OR spatial.col_min != e.dst_col_min
                OR spatial.col_max != e.dst_col_max
            """
        else:
            spatial_join = "LEFT JOIN main.edge_intervals AS spatial ON spatial.edge_id = e.id"
            mismatch = """
                spatial.edge_id IS NULL
                OR spatial.sheet_id != e.dst_sheet_id
                OR spatial.row_min != e.dst_row_min
                OR spatial.row_max != e.dst_row_max
                OR spatial.col_min != e.dst_col_min
                OR spatial.col_max != e.dst_col_max
            """
        row = self._connection.execute(
            f"""
            SELECT e.id
            FROM main.edges AS e
            {spatial_join}
            WHERE e.dst_sheet_id = ?
              AND e.dst_row_min <= ? AND e.dst_row_max >= ?
              AND e.dst_col_min <= ? AND e.dst_col_max >= ?
              AND ({mismatch})
            ORDER BY e.id
            LIMIT 1
            """,
            (
                area.sheet_id,
                area.rect.row_max,
                area.rect.row_min,
                area.rect.col_max,
                area.rect.col_min,
            ),
        ).fetchone()
        if row is not None:
            edge_id = _require_integral(row[0], 0, "edge id")
            _corrupt(edge_id, "relational destination does not match its spatial mirror")

    @staticmethod
    def _require_relational_rows(edge_ids: Sequence[int], rows: Sequence[dict[str, Any]]) -> None:
        relational_ids = {_require_integral(row["edge_id"], 0, "edge id") for row in rows}
        missing = next((edge_id for edge_id in edge_ids if edge_id not in relational_ids), None)
        if missing is not None:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                f"Spatial edge {missing} has no relational dependency edge.",
            )


_EDGE_SELECT = """
SELECT e.id AS edge_id, e.src_kind, e.src_id, e.src_sheet_id,
       e.dst_sheet_id, e.dst_row_min, e.dst_row_max, e.dst_col_min,
       e.dst_col_max, e.via, e.dependent_rank, e.precedent_rank,
       src_sheet.name AS src_sheet_name,
       fb.id AS fblock_id, fb.sheet_id AS fblock_sheet_id, fb.n AS fblock_n,
       fb.row_min AS src_row_min, fb.row_max AS src_row_max,
       fb.col_min AS src_col_min, fb.col_max AS src_col_max,
       dst_sheet.name AS dst_sheet_name
FROM main.edges AS e
LEFT JOIN main.sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
LEFT JOIN main.fblocks AS fb ON e.src_kind = 'fblock' AND fb.id = e.src_id
LEFT JOIN main.sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
"""


def _source_target(row: dict[str, Any]) -> GraphTarget:
    edge_id = _require_integral(row["edge_id"], 0, "edge id")
    source_sheet = row["src_sheet_name"]
    if source_sheet is None:
        _corrupt(edge_id, "source sheet does not exist")
    sheet_id = _require_integral(row["src_sheet_id"], edge_id, "source sheet id")
    sheet = str(source_sheet)
    source_kind = str(row["src_kind"])
    if source_kind == "fblock":
        if (
            row["fblock_id"] is None
            or _require_integral(row["fblock_sheet_id"], edge_id, "formula-block sheet id")
            != sheet_id
        ):
            _corrupt(edge_id, "formula-block source is orphaned or on another sheet")
        rect = _rect_from_values(
            edge_id,
            row["src_row_min"],
            row["src_row_max"],
            row["src_col_min"],
            row["src_col_max"],
            label="source",
        )
        area = GraphArea(sheet_id, sheet, rect)
        block_n = _require_integral(row["fblock_n"], edge_id, "formula-block ordinal")
        if block_n < 0:
            _corrupt(edge_id, "formula-block ordinal is negative")
        symbol = formula_block_symbol_id(sheet, block_n)
        return GraphTarget("fblock", symbol, area.ref, area)
    if source_kind == "cell":
        packed = _require_integral(row["src_id"], edge_id, "packed cell source")
        source_row, source_col = divmod(packed, _CELL_ID_FACTOR)
        if not 1 <= source_row <= _ROW_MAX or not 1 <= source_col <= _COL_MAX:
            _corrupt(edge_id, "packed cell source is outside worksheet bounds")
        area = GraphArea(sheet_id, sheet, Rect(source_row, source_row, source_col, source_col))
        return GraphTarget(
            "cell", cell_symbol_id(sheet, area.ref.rsplit("!", 1)[-1]), area.ref, area
        )
    _corrupt(edge_id, f"unsupported source kind {source_kind!r}")


def _destination_target(row: dict[str, Any]) -> GraphTarget:
    edge_id = _require_integral(row["edge_id"], 0, "edge id")
    values = (
        row["dst_sheet_id"],
        row["dst_row_min"],
        row["dst_row_max"],
        row["dst_col_min"],
        row["dst_col_max"],
    )
    if all(value is None for value in values):
        return GraphTarget("opaque", None, str(row["via"]), None)
    if any(value is None for value in values):
        _corrupt(edge_id, "destination is a partial rectangle")
    if row["dst_sheet_name"] is None:
        _corrupt(edge_id, "destination sheet does not exist")
    sheet_id = _require_integral(row["dst_sheet_id"], edge_id, "destination sheet id")
    sheet = str(row["dst_sheet_name"])
    rect = _rect_from_values(edge_id, *values[1:], label="destination")
    area = GraphArea(sheet_id, sheet, rect)
    if rect.row_min == rect.row_max and rect.col_min == rect.col_max:
        local_ref = area.ref.rsplit("!", 1)[-1]
        return GraphTarget("cell", cell_symbol_id(sheet, local_ref), area.ref, area)
    return GraphTarget("range", None, area.ref, area)


def _rect_from_values(
    edge_id: int,
    row_min: Any,
    row_max: Any,
    col_min: Any,
    col_max: Any,
    *,
    label: str,
) -> Rect:
    try:
        return Rect(
            _require_integral(row_min, edge_id, f"{label} row minimum"),
            _require_integral(row_max, edge_id, f"{label} row maximum"),
            _require_integral(col_min, edge_id, f"{label} column minimum"),
            _require_integral(col_max, edge_id, f"{label} column maximum"),
        )
    except (TypeError, ValueError) as exc:
        raise ExcelLSPError(
            ErrorCode.CORRUPT,
            f"Dependency edge {edge_id} has an invalid {label} rectangle.",
        ) from exc


def _corrupt(edge_id: int, problem: str) -> NoReturn:
    raise ExcelLSPError(
        ErrorCode.CORRUPT,
        f"Dependency edge {edge_id} is corrupt: {problem}.",
    )


def _require_integral(value: Any, edge_id: int, label: str) -> int:
    if type(value) is int:
        return value
    if type(value) is float and value.is_integer():
        return int(value)
    _corrupt(edge_id, f"{label} is not an integer")


def _require_area(value: GraphArea | GraphTarget) -> GraphArea:
    if isinstance(value, GraphArea):
        return value
    if value.area is None:
        raise ValueError("opaque graph targets cannot be expanded")
    return value.area


def _root_target(value: GraphArea | GraphTarget) -> GraphTarget:
    if isinstance(value, GraphTarget):
        return value
    if value.rect.row_min == value.rect.row_max and value.rect.col_min == value.rect.col_max:
        local_ref = value.ref.rsplit("!", 1)[-1]
        return GraphTarget("cell", cell_symbol_id(value.sheet, local_ref), value.ref, value)
    return GraphTarget("range", None, value.ref, value)


def _target_key(target: GraphTarget) -> _NodeKey:
    if target.area is None:
        return ("opaque", 0, 0, 0, 0, 0, target.ref or "")
    rect = target.area.rect
    return (
        target.kind,
        target.area.sheet_id,
        rect.row_min,
        rect.row_max,
        rect.col_min,
        rect.col_max,
        target.symbol or "",
    )


def _hop_sort_key(hop: GraphHop) -> tuple[object, ...]:
    target = hop.target
    if target.area is None:
        return (
            1,
            "",
            "",
            0,
            0,
            0,
            0,
            target.kind,
            target.label.casefold(),
            target.label,
            hop.via.casefold(),
            hop.via,
        )
    rect = target.area.rect
    return (
        0,
        target.area.sheet.casefold(),
        target.area.sheet,
        rect.row_min,
        rect.col_min,
        rect.row_max,
        rect.col_max,
        target.kind,
        target.label.casefold(),
        target.label,
        hop.via.casefold(),
        hop.via,
    )


def _deduplicate_hops(hops: Iterable[GraphHop]) -> tuple[GraphHop, ...]:
    unique: dict[tuple[_NodeKey, str], GraphHop] = {}
    for hop in hops:
        unique.setdefault((_target_key(hop.target), hop.via), hop)
    return tuple(sorted(unique.values(), key=_hop_sort_key))


def _freeze_trace(node: _MutableTraceNode) -> TraceNode:
    return TraceNode(
        node.target,
        node.via,
        tuple(_freeze_trace(child) for child in node.children),
        node.child_count,
    )


def _areas_intersect(left: GraphArea, right: GraphArea) -> bool:
    return left.sheet_id == right.sheet_id and left.rect.intersects(right.rect)


def _validate_bounded_int(name: str, value: object, *, minimum: int, maximum: int) -> None:
    if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= maximum:
        raise ValueError(f"{name} must be an integer between {minimum} and {maximum}")


def _reconstruct_paths(
    destination: _NodeKey,
    source: _NodeKey,
    parents: dict[_NodeKey, list[tuple[_NodeKey, str]]],
    limit: int,
) -> tuple[tuple[tuple[tuple[_NodeKey, str | None], ...], ...], bool]:
    paths: list[tuple[tuple[_NodeKey, str | None], ...]] = []

    def visit(node: _NodeKey, suffix: tuple[tuple[_NodeKey, str], ...]) -> None:
        if len(paths) > limit:
            return
        if node == source:
            paths.append(((source, None), *suffix))
            return
        for parent, via in sorted(parents[node]):
            visit(parent, ((node, via), *suffix))

    visit(destination, ())
    return tuple(paths[:limit]), len(paths) > limit


__all__ = ["DependencyGraph"]
