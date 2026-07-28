from __future__ import annotations

import sqlite3
from collections.abc import Sequence
from contextlib import suppress
from pathlib import Path
from threading import Event, Thread
from typing import Any, Self, cast

import pytest

import excel_lsp.core.index.edges as edge_module
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.exception_evidence import (
    normalize_exception_graph,
    prepare_chained_failure,
)
from excel_lsp.core.graph.models import GraphArea
from excel_lsp.core.graph.queries import DependencyGraph
from excel_lsp.core.index.edges import EdgeSchemaError, EdgeStore, SQLiteConnectionLike
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import Rect, SheetDescriptor

_SURFACES = (
    "direct_precedents",
    "direct_dependents",
    "trace_precedents",
    "trace_dependents",
    "trace_path",
)
_PROTECTED_TEMP_NAMES = (
    "edges",
    "fblocks",
    "sheets",
    "graph_spatial_state",
    "graph_rank_keys",
    "edge_rtree",
    "edge_source_rtree",
    "edge_intervals",
    "edge_source_intervals",
    *(
        f"{table}_graph_spatial_dirty_{operation}"
        for table in ("edges", "fblocks", "sheets", "graph_rank_keys")
        for operation in ("insert", "update", "delete")
    ),
    *(
        f"{table}_graph_dirty_{operation}"
        for table in (
            "edge_rtree_rowid",
            "edge_source_rtree_rowid",
            "edge_intervals",
            "edge_source_intervals",
        )
        for operation in ("insert", "update", "delete")
    ),
)


class _TaggedExceptionGroup(BaseExceptionGroup):
    tag: str

    def __new__(
        cls,
        message: str,
        exceptions: Sequence[BaseException],
        tag: str,
    ) -> Self:
        del tag
        return cast(Self, super().__new__(cls, message, exceptions))

    def __init__(
        self,
        message: str,
        exceptions: Sequence[BaseException],
        tag: str,
    ) -> None:
        super().__init__(message, exceptions)
        self.tag = tag

    def derive(
        self,
        exceptions: Sequence[BaseException],
    ) -> BaseExceptionGroup:
        return type(self)(self.message, exceptions, self.tag)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("surface", _SURFACES)
@pytest.mark.parametrize("rebuild_before_damage", [False, True], ids=["no-rebuild", "post-rebuild"])
def test_managed_commit_cannot_bless_coherently_restored_duplicate_rank_split(
    tmp_path: Path,
    prefer_rtree: bool,
    surface: str,
    rebuild_before_damage: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"managed-{surface}-{rebuild_before_damage}") as store:
        graph = store.dependency_graph
        assert _invoke_surface(graph, surface) is not None

        with store.transaction():
            if rebuild_before_damage:
                store.rebuild_graph_spatial_index()
            _split_duplicate_and_restore_persisted_seals(store, mode="execute")

        with pytest.raises(ExcelLSPError) as captured:
            _invoke_surface(graph, surface)
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("after_rebuild", [False, True], ids=["no-rebuild", "post-rebuild"])
def test_managed_non_graph_writes_do_not_invalidate_graph_seal(
    tmp_path: Path,
    prefer_rtree: bool,
    after_rebuild: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"non-graph-{after_rebuild}") as store:
        before = tuple(_invoke_surface(store.dependency_graph, surface) for surface in _SURFACES)
        with store.transaction():
            if after_rebuild:
                store.rebuild_graph_spatial_index()
            store.set_meta("p4_non_graph_probe", "accepted")
            store.connection.execute(
                "INSERT INTO diagnostics(severity, code, sheet_id, ref, message, related) "
                "VALUES ('info', 'I_PROBE', 1, 'Inputs!A1', 'probe', '{}')"
            )

        assert store.get_meta("p4_non_graph_probe") == "accepted"
        after = tuple(_invoke_surface(store.dependency_graph, surface) for surface in _SURFACES)
        assert after == before


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("surface", _SURFACES)
def test_public_connection_aliases_cannot_displace_tracking(
    tmp_path: Path,
    prefer_rtree: bool,
    surface: str,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"capability-alias-{surface}") as store:
        graph = store.dependency_graph
        assert _invoke_surface(graph, surface) is not None
        connection = store.connection
        cursor = connection.cursor()
        execute_cursor = connection.execute("SELECT 1")
        try:
            with connection as entered:
                aliases = (
                    connection,
                    cursor.connection,
                    execute_cursor.connection,
                    entered,
                )
                assert all(alias is connection for alias in aliases)
                assert all(not isinstance(alias, sqlite3.Connection) for alias in aliases)
                for alias in aliases:
                    with pytest.raises(TypeError):
                        sqlite3.Connection.set_authorizer(cast(Any, alias), None)
        finally:
            cursor.close()
            execute_cursor.close()

        with store.transaction():
            _split_duplicate_and_restore_persisted_seals(store, mode="execute")

        with pytest.raises(ExcelLSPError) as captured:
            _invoke_surface(graph, surface)
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_public_cursor_facade_preserves_protocol_without_native_leaks(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "cursor-capability") as store:
        connection = store.connection
        assert not isinstance(connection, sqlite3.Connection)

        with pytest.raises(TypeError, match="custom cursor factories"):
            connection.cursor(object())

        def row_factory(cursor: sqlite3.Cursor, row: tuple[object, ...]) -> tuple[object, ...]:
            del cursor
            return row

        with pytest.raises(TypeError, match="custom connection row factories"):
            connection.row_factory = row_factory
        assert connection.row_factory is sqlite3.Row

        cursor = connection.cursor()
        try:
            assert not isinstance(cursor, sqlite3.Cursor)
            assert cursor.connection is connection
            descriptor = cast(Any, sqlite3.Cursor.connection)
            with pytest.raises(TypeError):
                descriptor.__get__(cursor, sqlite3.Cursor)
            with pytest.raises(TypeError, match="custom cursor row factories"):
                cursor.row_factory = row_factory
            assert cursor.row_factory is None

            assert cursor.execute("SELECT ?", (17,)) is cursor
            assert tuple(cursor.fetchone() or ()) == (17,)
            assert (
                cursor.executemany(
                    "INSERT INTO diagnostics("
                    "severity, code, sheet_id, ref, message, related"
                    ") VALUES ('info', ?, 1, 'Inputs!A1', 'probe', '{}')",
                    (("I_CURSOR_ONE",), ("I_CURSOR_TWO",)),
                )
                is cursor
            )
            assert cursor.executescript("SELECT 1;") is cursor
            assert cursor.connection is connection
        finally:
            cursor.close()

        execute_cursor = connection.execute("SELECT 1")
        try:
            assert not isinstance(execute_cursor, sqlite3.Cursor)
            assert execute_cursor.connection is connection
        finally:
            execute_cursor.close()

        executemany_cursor = connection.executemany(
            "INSERT INTO diagnostics("
            "severity, code, sheet_id, ref, message, related"
            ") VALUES ('info', ?, 1, 'Inputs!A1', 'probe', '{}')",
            (("I_CONNECTION_MANY",),),
        )
        try:
            assert not isinstance(executemany_cursor, sqlite3.Cursor)
            assert executemany_cursor.connection is connection
        finally:
            executemany_cursor.close()

        script_cursor = connection.executescript("SELECT 1;")
        try:
            assert not isinstance(script_cursor, sqlite3.Cursor)
            assert script_cursor.connection is connection
        finally:
            script_cursor.close()

        for surface in _SURFACES:
            assert _invoke_surface(store.dependency_graph, surface) is not None


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_graph_mutation_before_final_rebuild_is_sealed_and_accepted(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "mutation-before-rebuild") as store:
        with store.transaction():
            store.connection.execute("UPDATE edges SET via = 'split-before-rebuild' WHERE id = 3")
            store.rebuild_graph_spatial_index()
            store.set_meta("after_graph_rebuild", "non-graph")

        for surface in _SURFACES:
            assert _invoke_surface(store.dependency_graph, surface) is not None


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_rollback_rebases_monotonic_graph_evidence_only_after_state_restoration(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "rollback") as store:
        before = tuple(_invoke_surface(store.dependency_graph, surface) for surface in _SURFACES)
        with pytest.raises(RuntimeError, match="force graph rollback"), store.transaction():
            store.rebuild_graph_spatial_index()
            _split_duplicate_and_restore_persisted_seals(store, mode="execute")
            raise RuntimeError("force graph rollback")

        after = tuple(_invoke_surface(store.dependency_graph, surface) for surface in _SURFACES)
        assert after == before


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("mode", ["execute", "cursor", "executemany"])
def test_same_handle_epoch_observes_every_sql_execution_surface(
    tmp_path: Path,
    prefer_rtree: bool,
    mode: str,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"execution-{mode}") as store:
        graph = store.dependency_graph
        # Execute the same text once before the adversarial transaction. A tracker
        # that only runs when SQLite first prepares a cached statement is vacuous.
        with store.transaction():
            _mutate_duplicate(store.connection, mode, via="temporary")
            store.rebuild_graph_spatial_index()

        with store.transaction():
            _split_duplicate_and_restore_persisted_seals(store, mode=mode)

        with pytest.raises(ExcelLSPError) as captured:
            graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_client_authorizer_is_chained_without_displacing_graph_tracking(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "authorizer-chain") as store:
        observed: list[tuple[int, str | None]] = []

        def observe(
            action: int,
            argument_one: str | None,
            argument_two: str | None,
            database_name: str | None,
            trigger: str | None,
        ) -> int:
            del argument_two, database_name, trigger
            observed.append((action, argument_one))
            return sqlite3.SQLITE_OK

        store.connection.set_authorizer(observe)
        try:
            with store.transaction():
                _split_duplicate_and_restore_persisted_seals(store, mode="execute")
        finally:
            store.connection.set_authorizer(None)

        assert any(
            action == sqlite3.SQLITE_UPDATE and table == "edges" for action, table in observed
        )
        with pytest.raises(ExcelLSPError) as captured:
            store.dependency_graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_client_authorizer_denial_remains_effective_and_does_not_poison_graph(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "authorizer-deny") as store:

        def deny_meta_update(
            action: int,
            argument_one: str | None,
            argument_two: str | None,
            database_name: str | None,
            trigger: str | None,
        ) -> int:
            del argument_two, database_name, trigger
            if action == sqlite3.SQLITE_UPDATE and argument_one == "meta":
                return sqlite3.SQLITE_DENY
            return sqlite3.SQLITE_OK

        store.connection.set_authorizer(deny_meta_update)
        try:
            with pytest.raises(sqlite3.DatabaseError):
                store.set_meta("generation", store.generation + 1)
        finally:
            store.connection.set_authorizer(None)

        for surface in _SURFACES:
            assert _invoke_surface(store.dependency_graph, surface) is not None


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_client_authorizer_denied_edge_update_changes_no_graph_authority(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "authorizer-deny-edge") as store:
        before_edges = _edge_rows(store.connection)
        before_state = _graph_state(store.connection)
        before_catalog = _catalog(store.connection)

        def deny_edge_update(
            action: int,
            argument_one: str | None,
            argument_two: str | None,
            database_name: str | None,
            trigger: str | None,
        ) -> int:
            del argument_two, database_name, trigger
            if action == sqlite3.SQLITE_UPDATE and argument_one == "edges":
                return sqlite3.SQLITE_DENY
            return sqlite3.SQLITE_OK

        store.connection.set_authorizer(deny_edge_update)
        try:
            with pytest.raises(sqlite3.DatabaseError):
                store.connection.execute("UPDATE edges SET via = 'denied' WHERE id = 3")
        finally:
            store.connection.set_authorizer(None)

        assert not store.connection.in_transaction
        assert _edge_rows(store.connection) == before_edges
        assert _graph_state(store.connection) == before_state
        assert _catalog(store.connection) == before_catalog
        for surface in _SURFACES:
            assert _invoke_surface(store.dependency_graph, surface) is not None


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_client_authorizer_ignore_write_invalidates_every_cached_graph_surface(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "authorizer-ignore-edge") as store:
        graph = store.dependency_graph
        for surface in _SURFACES:
            assert _invoke_surface(graph, surface) is not None
        before_epoch = _graph_write_epoch(store.connection)
        observed: list[tuple[int, str | None]] = []

        def ignore_delete_and_update(
            action: int,
            argument_one: str | None,
            argument_two: str | None,
            database_name: str | None,
            trigger: str | None,
        ) -> int:
            del argument_two, database_name, trigger
            observed.append((action, argument_one))
            if action in {sqlite3.SQLITE_DELETE, sqlite3.SQLITE_UPDATE}:
                return sqlite3.SQLITE_IGNORE
            return sqlite3.SQLITE_OK

        store.connection.set_authorizer(ignore_delete_and_update)
        try:
            store.connection.execute("DELETE FROM edges WHERE id = 1")
        finally:
            store.connection.set_authorizer(None)

        assert any(
            action == sqlite3.SQLITE_DELETE and table == "edges" for action, table in observed
        )
        assert (
            store.connection.execute("SELECT COUNT(*) FROM edges WHERE id = 1").fetchone()[0] == 0
        )
        assert _graph_write_epoch(store.connection) > before_epoch
        for surface in _SURFACES:
            with pytest.raises(ExcelLSPError) as captured:
                _invoke_surface(graph, surface)
            assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_public_authorizer_policy_removal_preserves_internal_tracking(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "authorizer-removal") as store:
        observed: list[tuple[int, str | None]] = []

        def deny_meta_update(
            action: int,
            argument_one: str | None,
            argument_two: str | None,
            database_name: str | None,
            trigger: str | None,
        ) -> int:
            del argument_two, database_name, trigger
            observed.append((action, argument_one))
            if action == sqlite3.SQLITE_UPDATE and argument_one == "meta":
                return sqlite3.SQLITE_DENY
            return sqlite3.SQLITE_OK

        store.connection.set_authorizer(deny_meta_update)
        store.connection.execute("SELECT value FROM meta WHERE key = 'generation'").fetchone()
        with pytest.raises(sqlite3.DatabaseError):
            store.set_meta("p4_authorizer_probe", "denied")
        assert any(action == sqlite3.SQLITE_READ and table == "meta" for action, table in observed)
        assert any(
            action == sqlite3.SQLITE_UPDATE and table == "meta" for action, table in observed
        )

        store.connection.set_authorizer(None)
        store.set_meta("p4_authorizer_probe", "accepted")
        assert store.get_meta("p4_authorizer_probe") == "accepted"
        for surface in _SURFACES:
            assert _invoke_surface(store.dependency_graph, surface) is not None

        with store.transaction():
            _split_duplicate_and_restore_persisted_seals(store, mode="execute")

        for surface in _SURFACES:
            with pytest.raises(ExcelLSPError) as captured:
                _invoke_surface(store.dependency_graph, surface)
            assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("surface", _SURFACES)
def test_temp_graph_shadow_is_rejected_and_trips_every_live_graph_surface(
    tmp_path: Path,
    prefer_rtree: bool,
    surface: str,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"temp-shadow-{surface}") as store:
        graph = store.dependency_graph
        assert _invoke_surface(graph, surface) is not None

        store.connection.execute("CREATE TEMP TABLE edges(probe INTEGER)")

        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            EdgeStore(store.connection)
        with pytest.raises(ExcelLSPError) as captured:
            _invoke_surface(graph, surface)
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("surface", _SURFACES)
def test_plain_connection_temp_ddl_trips_the_constant_size_live_seal(
    tmp_path: Path,
    prefer_rtree: bool,
    surface: str,
) -> None:
    store = _seed_store(tmp_path, prefer_rtree, f"plain-temp-shadow-{surface}")
    database = store.path
    store.close()
    connection = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    connection.row_factory = sqlite3.Row
    try:
        edges = EdgeStore(connection)
        graph = DependencyGraph(connection, edges)
        assert _invoke_surface(graph, surface) is not None
        changes_before = connection.total_changes

        connection.execute("CREATE TEMP TABLE edges(probe INTEGER)")

        assert connection.total_changes == changes_before
        with pytest.raises(ExcelLSPError) as captured:
            _invoke_surface(graph, surface)
        assert captured.value.code is ErrorCode.CORRUPT
    finally:
        connection.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_canonical_export_fails_closed_after_a_protected_temp_shadow(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "canonical-export-temp-shadow") as store:
        assert len(store.canonical_export()["edges"]) == 3
        store.connection.execute(
            "CREATE TEMP TABLE edges("
            "id, src_kind, src_id, src_sheet_id, dst_sheet_id, "
            "dst_row_min, dst_row_max, dst_col_min, dst_col_max, via, "
            "dependent_rank, precedent_rank)"
        )

        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            store.canonical_export()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    "temp_ddl",
    [
        "CREATE TEMP VIEW EdGeS AS SELECT 1 AS probe",
        ("CREATE TEMP TRIGGER unrelated_name AFTER UPDATE ON main.edges BEGIN SELECT 1; END"),
        "CREATE TEMP TABLE edge_rtree_node(probe INTEGER)",
        "CREATE TEMP TABLE edge_source_rtree_parent(probe INTEGER)",
    ],
    ids=["mixed-case-view", "protected-target", "rtree-prefix", "source-rtree-prefix"],
)
def test_every_protected_temp_object_shape_fails_closed(
    tmp_path: Path,
    prefer_rtree: bool,
    temp_ddl: str,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "temp-object-shapes") as store:
        graph = store.dependency_graph
        for surface in _SURFACES:
            assert _invoke_surface(graph, surface) is not None

        store.connection.execute(temp_ddl)

        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            EdgeStore(store.connection)
        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            store.canonical_export()
        for surface in _SURFACES:
            with pytest.raises(ExcelLSPError) as captured:
                _invoke_surface(graph, surface)
            assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("protected_name", _PROTECTED_TEMP_NAMES)
def test_every_exact_protected_temp_name_is_rejected(
    tmp_path: Path,
    prefer_rtree: bool,
    protected_name: str,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"temp-name-{protected_name}") as store:
        store.connection.execute(f'CREATE TEMP VIEW "{protected_name}" AS SELECT 1 AS probe')

        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            EdgeStore(store.connection)
        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            store.canonical_export()


@pytest.mark.parametrize(
    "target",
    ["edge_rtree_node", "edge_source_rtree_parent"],
    ids=["destination-shadow-target", "source-shadow-target"],
)
def test_arbitrary_temp_trigger_targeting_an_rtree_shadow_is_rejected(
    tmp_path: Path,
    target: str,
) -> None:
    with _seed_store(tmp_path, True, f"temp-target-{target}") as store:
        graph = store.dependency_graph
        for surface in _SURFACES:
            assert _invoke_surface(graph, surface) is not None
        store.connection.execute(
            f"CREATE TEMP TRIGGER unrelated_{target} AFTER UPDATE ON main.{target} "
            "BEGIN SELECT 1; END"
        )

        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            EdgeStore(store.connection)
        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            store.canonical_export()
        for surface in _SURFACES:
            with pytest.raises(ExcelLSPError) as captured:
                _invoke_surface(graph, surface)
            assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    "object_name",
    ["graph_spatial_state", "edge_rtree_probe", "edge_source_rtree_probe"],
    ids=["protected-name", "rtree-prefix-name", "source-rtree-prefix-name"],
)
def test_protected_temp_names_are_rejected_with_an_unprotected_target(
    tmp_path: Path,
    prefer_rtree: bool,
    object_name: str,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"temp-name-only-{object_name}") as store:
        store.connection.execute("CREATE TEMP TABLE safe_target(probe INTEGER)")
        store.connection.execute(
            f'CREATE TEMP TRIGGER "{object_name}" AFTER UPDATE ON safe_target BEGIN SELECT 1; END'
        )

        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            EdgeStore(store.connection)
        with pytest.raises(EdgeSchemaError, match="shadows graph storage"):
            store.canonical_export()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_rolled_back_temp_ddl_restores_every_surface_without_epoch_poisoning(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "rolled-back-temp-ddl") as store:
        graph = store.dependency_graph
        for surface in _SURFACES:
            assert _invoke_surface(graph, surface) is not None

        marker = RuntimeError("force TEMP DDL rollback")
        with pytest.raises(RuntimeError) as captured, store.transaction():
            store.connection.execute("CREATE TEMP TABLE edges(probe INTEGER)")
            raise marker
        assert captured.value is marker
        assert (
            store.connection.execute(
                "SELECT 1 FROM temp.sqlite_master WHERE name = 'edges'"
            ).fetchone()
            is None
        )
        epoch_after_rollback = _graph_write_epoch(store.connection)

        for surface in _SURFACES:
            assert _invoke_surface(graph, surface) is not None
        assert _graph_write_epoch(store.connection) == epoch_after_rollback


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("surface", _SURFACES)
@pytest.mark.parametrize("mode", ["execute", "cursor", "executemany"])
def test_writable_schema_trigger_mutation_advances_trust_boundary(
    tmp_path: Path,
    prefer_rtree: bool,
    surface: str,
    mode: str,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"writable-schema-{mode}-{surface}") as store:
        graph = store.dependency_graph
        assert _invoke_surface(graph, surface) is not None
        trigger_name = "edges_graph_spatial_dirty_update"
        original_row = store.connection.execute(
            "SELECT sql FROM main.sqlite_master WHERE type = 'trigger' AND name = ?",
            (trigger_name,),
        ).fetchone()
        assert original_row is not None and isinstance(original_row[0], str)

        store.connection.execute("PRAGMA writable_schema = ON")
        epoch_before_catalog_write = _graph_write_epoch(store.connection)
        try:
            sql = (
                "UPDATE main.sqlite_master SET sql = sql || ' ' WHERE type = 'trigger' AND name = ?"
            )
            if mode == "execute":
                store.connection.execute(sql, (trigger_name,))
            elif mode == "cursor":
                cursor = store.connection.cursor()
                try:
                    cursor.execute(sql, (trigger_name,))
                finally:
                    cursor.close()
            elif mode == "executemany":
                store.connection.executemany(sql, ((trigger_name,),))
            else:
                raise AssertionError(f"unknown execution mode {mode!r}")
        finally:
            store.connection.execute("PRAGMA writable_schema = OFF")

        assert _graph_write_epoch(store.connection) > epoch_before_catalog_write
        changed_row = store.connection.execute(
            "SELECT sql FROM main.sqlite_master WHERE type = 'trigger' AND name = ?",
            (trigger_name,),
        ).fetchone()
        assert changed_row is not None and changed_row[0] == f"{original_row[0]} "
        with pytest.raises(ExcelLSPError) as captured:
            _invoke_surface(graph, surface)
        assert captured.value.code is ErrorCode.CORRUPT


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
def test_constructor_validation_and_live_seal_share_one_writer_stable_snapshot(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    journal_mode: str,
) -> None:
    store = _seed_store(tmp_path, prefer_rtree, f"constructor-{journal_mode}")
    database = store.path
    state = _graph_state(store.connection)
    catalog = _catalog(store.connection)
    store.close()

    reader = sqlite3.connect(database, timeout=5.0, isolation_level=None)
    reader.row_factory = sqlite3.Row
    assert (
        str(reader.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()[0]) == journal_mode
    )
    reader.execute("PRAGMA busy_timeout = 5000")

    writer_ready = Event()
    validation_complete = Event()
    writer_attempted = Event()
    writer_committed = Event()
    writer_errors: list[BaseException] = []

    def writer() -> None:
        connection = sqlite3.connect(database, timeout=5.0, isolation_level=None)
        try:
            connection.execute("PRAGMA busy_timeout = 5000")
            assert str(connection.execute("PRAGMA journal_mode").fetchone()[0]) == journal_mode
            writer_ready.set()
            if not validation_complete.wait(5):
                raise TimeoutError("constructor validation never reached its barrier")
            writer_attempted.set()
            connection.execute("BEGIN IMMEDIATE")
            connection.execute("UPDATE edges SET via = 'constructor-split' WHERE id = 3")
            connection.executemany(
                "INSERT OR REPLACE INTO graph_rank_keys(direction, rank, key_text) "
                "VALUES (?, ?, ?)",
                catalog,
            )
            _restore_graph_state(connection, state)
            connection.commit()
            writer_committed.set()
        except BaseException as error:
            writer_errors.append(error)
        finally:
            writer_ready.set()
            writer_attempted.set()
            writer_committed.set()
            connection.close()

    committed_during_validation: list[bool] = []
    original_validate = edge_module._validate_graph_sidecar

    def validate_with_barrier(
        connection: sqlite3.Connection,
        backend: Any,
        tables: set[str],
    ) -> None:
        original_validate(connection, backend, tables)
        validation_complete.set()
        assert writer_attempted.wait(5)
        committed_during_validation.append(writer_committed.wait(0.25))

    thread = Thread(target=writer, daemon=True)
    thread.start()
    assert writer_ready.wait(5)
    try:
        with monkeypatch.context() as patch:
            patch.setattr(edge_module, "_validate_graph_sidecar", validate_with_barrier)
            edges = EdgeStore(reader)
        thread.join(10)
        assert not thread.is_alive()
        assert writer_errors == []
        assert committed_during_validation == [False]

        graph = DependencyGraph(reader, edges)
        with pytest.raises(ExcelLSPError) as captured:
            graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
        assert captured.value.code is ErrorCode.CORRUPT
    finally:
        validation_complete.set()
        thread.join(10)
        reader.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("failure_timing", ["before", "after"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
def test_constructor_begin_immediate_failure_releases_any_acquired_writer_lock(
    tmp_path: Path,
    prefer_rtree: bool,
    failure_timing: str,
    journal_mode: str,
) -> None:
    store = _seed_store(
        tmp_path,
        prefer_rtree,
        f"begin-failure-{journal_mode}-{failure_timing}",
    )
    database = store.path
    store.close()

    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_BeginImmediateFailureConnection,
    )
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA busy_timeout = 1000")
    assert (
        str(connection.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()[0])
        == journal_mode
    )
    assert str(connection.execute("PRAGMA journal_mode").fetchone()[0]) == journal_mode
    marker = sqlite3.DatabaseError(f"BEGIN IMMEDIATE failed {failure_timing} native effect")
    connection.begin_failure_timing = failure_timing
    connection.begin_failure_marker = marker

    try:
        with pytest.raises(sqlite3.DatabaseError) as captured:
            EdgeStore(connection)
        assert captured.value is marker
        _assert_connection_is_inactive_or_closed(connection)
        _assert_writer_can_begin(database)
    finally:
        with suppress(sqlite3.ProgrammingError):
            connection.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("begin_sql", ["BEGIN", "BEGIN IMMEDIATE"], ids=["deferred", "immediate"])
def test_constructor_rejects_without_disturbing_caller_owned_transaction(
    tmp_path: Path,
    prefer_rtree: bool,
    begin_sql: str,
) -> None:
    store = _seed_store(tmp_path, prefer_rtree, f"caller-transaction-{begin_sql}")
    database = store.path
    store.close()

    connection = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    connection.row_factory = sqlite3.Row
    try:
        connection.execute(begin_sql)
        assert connection.in_transaction
        assert int(connection.execute("SELECT COUNT(*) FROM edges").fetchone()[0]) == 3

        with pytest.raises(RuntimeError, match="outside a transaction"):
            EdgeStore(connection)

        assert connection.in_transaction
        connection.execute(
            "INSERT INTO meta(key, value) VALUES ('caller_transaction_probe', 'uncommitted')"
        )
        assert (
            connection.execute(
                "SELECT value FROM meta WHERE key = 'caller_transaction_probe'"
            ).fetchone()[0]
            == "uncommitted"
        )
        connection.rollback()
        assert not connection.in_transaction
        assert (
            connection.execute(
                "SELECT value FROM meta WHERE key = 'caller_transaction_probe'"
            ).fetchone()
            is None
        )
        _assert_writer_can_begin(database)
    finally:
        if connection.in_transaction:
            connection.rollback()
        connection.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("journal_mode", ["wal", "delete"])
def test_constructor_waits_for_existing_immediate_writer_before_validation(
    tmp_path: Path,
    prefer_rtree: bool,
    journal_mode: str,
) -> None:
    store = _seed_store(tmp_path, prefer_rtree, f"writer-first-{journal_mode}")
    database = store.path
    state = _graph_state(store.connection)
    catalog = _catalog(store.connection)
    store.close()

    writer = sqlite3.connect(database, timeout=5.0, isolation_level=None)
    writer.row_factory = sqlite3.Row
    assert (
        str(writer.execute(f"PRAGMA journal_mode = {journal_mode}").fetchone()[0]) == journal_mode
    )
    writer.execute("PRAGMA busy_timeout = 5000")
    writer.execute("BEGIN IMMEDIATE")
    writer.execute("UPDATE edges SET via = 'writer-owned-split' WHERE id = 3")
    writer.executemany(
        "INSERT OR REPLACE INTO graph_rank_keys(direction, rank, key_text) VALUES (?, ?, ?)",
        catalog,
    )
    _restore_graph_state(writer, state)

    constructor_started = Event()
    constructor_finished = Event()
    constructor_errors: list[BaseException] = []

    def construct() -> None:
        connection = sqlite3.connect(database, timeout=5.0, isolation_level=None)
        connection.row_factory = sqlite3.Row
        try:
            connection.execute("PRAGMA busy_timeout = 5000")
            assert str(connection.execute("PRAGMA journal_mode").fetchone()[0]) == journal_mode
            constructor_started.set()
            EdgeStore(connection)
        except BaseException as error:
            constructor_errors.append(error)
        finally:
            constructor_started.set()
            constructor_finished.set()
            connection.close()

    thread = Thread(target=construct, daemon=True)
    thread.start()
    try:
        assert constructor_started.wait(5)
        assert not constructor_finished.wait(0.25)
        writer.commit()
        thread.join(10)
        assert not thread.is_alive()
        assert len(constructor_errors) == 1
        assert isinstance(constructor_errors[0], EdgeSchemaError)
        _assert_writer_can_begin(database)
    finally:
        if writer.in_transaction:
            writer.rollback()
        writer.close()
        thread.join(10)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("boundary", ["graph", "store"])
def test_nested_primary_membership_is_removed_before_cleanup_chaining(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    boundary: str,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"nested-primary-{boundary}") as store:
        primary = RuntimeError(f"{boundary} cleanup primary")
        other = RuntimeError(f"{boundary} distinct cleanup evidence")
        inner = BaseExceptionGroup("inner cleanup", (primary, other))
        outer = BaseExceptionGroup("outer cleanup", (inner,))

        if boundary == "graph":
            graph = store.dependency_graph
            original_release = graph._release_owned_read_snapshot

            def release_with_cycle_candidate() -> tuple[BaseException, ...]:
                assert original_release() == ()
                return (primary, outer)

            with monkeypatch.context() as patch:
                patch.setattr(graph, "_release_owned_read_snapshot", release_with_cycle_candidate)
                with pytest.raises(RuntimeError) as captured:
                    graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
        else:
            original_finalizer = store.edge_store.finalize_transaction_commit

            def commit_then_fail() -> None:
                original_finalizer()
                raise primary

            def finalize_then_fail() -> None:
                original_finalizer()
                raise outer

            with monkeypatch.context() as patch:
                patch.setattr(store.edge_store, "transaction_committed", commit_then_fail)
                patch.setattr(store.edge_store, "finalize_transaction_commit", finalize_then_fail)
                with pytest.raises(RuntimeError) as captured, store.transaction():
                    store.set_meta("nested_primary_boundary", boundary)

        assert captured.value is primary
        _assert_unique_exception_identities(captured.value)
        assert _exception_graph_contains(captured.value, other)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("boundary", ["graph", "store"])
def test_shared_external_causal_identity_has_one_owner_and_keeps_its_chain(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    boundary: str,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, f"shared-external-{boundary}") as store:
        primary = RuntimeError(f"{boundary} primary")
        descendant = OSError(f"{boundary} external descendant")
        shared = LookupError(f"{boundary} shared external")
        shared.__cause__ = descendant
        first = RuntimeError(f"{boundary} first cleanup")
        second = RuntimeError(f"{boundary} second cleanup")
        first.__cause__ = shared
        second.__context__ = shared
        group = BaseExceptionGroup("shared external cleanup", (first, second))

        if boundary == "graph":
            graph = store.dependency_graph
            original_release = graph._release_owned_read_snapshot

            def release_with_shared_external() -> tuple[BaseException, ...]:
                assert original_release() == ()
                return (primary, group)

            with monkeypatch.context() as patch:
                patch.setattr(graph, "_release_owned_read_snapshot", release_with_shared_external)
                with pytest.raises(RuntimeError) as captured:
                    graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
        else:
            original_finalizer = store.edge_store.finalize_transaction_commit

            def commit_then_fail() -> None:
                original_finalizer()
                raise primary

            def finalize_then_fail() -> None:
                original_finalizer()
                raise group

            with monkeypatch.context() as patch:
                patch.setattr(store.edge_store, "transaction_committed", commit_then_fail)
                patch.setattr(store.edge_store, "finalize_transaction_commit", finalize_then_fail)
                with pytest.raises(RuntimeError) as captured, store.transaction():
                    store.set_meta("shared_external_boundary", boundary)

        assert captured.value is primary
        _assert_unique_exception_identities(captured.value)
        assert _count_exception_identity(captured.value, shared) == 1
        assert _count_exception_identity(captured.value, descendant) == 1


def test_standalone_exception_sanitizer_preserves_group_metadata_and_suppression() -> None:
    shared = ValueError("shared membership")
    cycle_member = RuntimeError("cycle membership")
    cycle_external = OSError("cycle external")
    cycle_member.__cause__ = cycle_external
    cycle_external.__context__ = cycle_member

    nested = BaseExceptionGroup("nested metadata", (shared, cycle_member))
    nested.add_note("nested note")
    nested.__cause__ = shared
    nested.__suppress_context__ = False

    root_cause = LookupError("root cause")
    root_context = ArithmeticError("root context")
    root = BaseExceptionGroup("root metadata", (shared, nested, shared, nested))
    root.add_note("root note")
    root.__context__ = root_context
    root.__cause__ = root_cause
    root.__suppress_context__ = False
    try:
        raise root
    except BaseExceptionGroup as raised_root:
        root = raised_root

    source_traceback = root.__traceback__
    normalized = normalize_exception_graph(root)

    assert isinstance(normalized, BaseExceptionGroup)
    assert normalized is not root
    assert normalized.message == "root metadata"
    assert normalized.__notes__ == ["root note"]
    assert normalized.__cause__ is root_cause
    assert normalized.__context__ is root_context
    assert normalized.__suppress_context__ is False
    assert normalized.__traceback__ is source_traceback
    assert len(normalized.exceptions) == 2
    normalized_nested = normalized.exceptions[1]
    assert isinstance(normalized_nested, BaseExceptionGroup)
    assert normalized_nested.message == "nested metadata"
    assert normalized_nested.__notes__ == ["nested note"]
    assert normalized_nested.__cause__ is None
    assert normalized_nested.__suppress_context__ is False
    assert normalized_nested.exceptions == (cycle_member,)
    assert cycle_member.__cause__ is cycle_external
    assert cycle_external.__context__ is None
    _assert_unique_exception_identities(normalized)
    assert _count_exception_identity(normalized, shared) == 1
    assert _count_exception_identity(normalized, cycle_member) == 1
    assert _count_exception_identity(normalized, cycle_external) == 1


def test_exception_sanitizer_uses_subclass_derive_and_preserves_custom_state() -> None:
    shared = ValueError("tagged shared member")
    distinct = RuntimeError("tagged distinct member")
    source = _TaggedExceptionGroup(
        "tagged group",
        (shared, distinct, shared),
        "custom-state",
    )
    source.add_note("tagged note")
    source.__suppress_context__ = False
    try:
        raise source
    except _TaggedExceptionGroup as raised_source:
        source = raised_source
    source_traceback = source.__traceback__

    normalized = normalize_exception_graph(source)

    assert isinstance(normalized, _TaggedExceptionGroup)
    assert normalized is not source
    assert normalized.message == "tagged group"
    assert normalized.tag == "custom-state"
    assert normalized.exceptions == (shared, distinct)
    assert normalized.__notes__ == ["tagged note"]
    assert normalized.__traceback__ is source_traceback
    assert normalized.__suppress_context__ is False


def test_chained_sanitizer_excludes_the_complete_primary_exception_graph() -> None:
    member = ValueError("primary member")
    descendant = KeyError("primary member descendant")
    member.__cause__ = descendant
    descendant.__context__ = member
    cause = LookupError("primary cause")
    context = ArithmeticError("primary context")
    primary = BaseExceptionGroup("primary graph", (member,))
    primary.__cause__ = cause
    primary.__context__ = context
    primary.__suppress_context__ = False
    distinct = RuntimeError("distinct cleanup")
    cleanup = BaseExceptionGroup(
        "cleanup graph",
        (member, descendant, cause, context, distinct),
    )

    prepared = prepare_chained_failure(cleanup, primary)

    assert isinstance(prepared, BaseExceptionGroup)
    assert prepared is not cleanup
    assert prepared.exceptions == (distinct,)
    assert not _exception_graph_contains(prepared, member)
    assert not _exception_graph_contains(prepared, descendant)
    assert not _exception_graph_contains(prepared, cause)
    assert not _exception_graph_contains(prepared, context)


def test_normalizer_preserves_external_group_sibling_as_one_causal_descendant() -> None:
    first = RuntimeError("external first")
    second = ValueError("external second")
    first.__cause__ = second
    external_group = BaseExceptionGroup("external group", (first, second))
    root = OSError("cleanup root")
    root.__cause__ = external_group

    normalized = normalize_exception_graph(root)

    assert normalized is not None
    assert normalized is root
    assert isinstance(normalized.__cause__, BaseExceptionGroup)
    assert normalized.__cause__ is not external_group
    assert normalized.__cause__.exceptions == (first,)
    assert first.__cause__ is second
    _assert_unique_exception_identities(normalized)
    assert _count_exception_identity(normalized, first) == 1
    assert _count_exception_identity(normalized, second) == 1


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_graph_cleanup_boundary_preserves_external_group_sibling_evidence(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "external-group-boundary") as store:
        graph = store.dependency_graph
        first = RuntimeError("boundary external first")
        second = ValueError("boundary external second")
        first.__cause__ = second
        external_group = BaseExceptionGroup("boundary external group", (first, second))
        root = OSError("boundary cleanup root")
        root.__cause__ = external_group
        original_release = graph._release_owned_read_snapshot

        def release_with_external_group() -> tuple[BaseException, ...]:
            assert original_release() == ()
            return (root,)

        with monkeypatch.context() as patch:
            patch.setattr(graph, "_release_owned_read_snapshot", release_with_external_group)
            with pytest.raises(OSError) as captured:
                _invoke_surface(graph, "direct_dependents")

        assert captured.value is root
        _assert_unique_exception_identities(captured.value)
        assert _count_exception_identity(captured.value, first) == 1
        assert _count_exception_identity(captured.value, second) == 1


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("prior_link", ["cause", "context"])
@pytest.mark.parametrize("boundary", ["graph-cleanup", "post-commit", "constructor"])
def test_cleanup_composition_preserves_preexisting_primary_causal_evidence(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    prior_link: str,
    boundary: str,
) -> None:
    store = _seed_store(tmp_path, prefer_rtree, f"prior-evidence-{boundary}-{prior_link}")
    primary = RuntimeError(f"{boundary} primary")
    prior = ValueError(f"{boundary} prior {prior_link}")
    cleanup = OSError(f"{boundary} cleanup")
    if prior_link == "cause":
        primary.__cause__ = prior
    else:
        primary.__context__ = prior
        primary.__suppress_context__ = False
    connection: sqlite3.Connection | None = None
    try:
        if boundary == "graph-cleanup":
            graph = store.dependency_graph
            original_release = graph._release_owned_read_snapshot

            def release_graph_failures() -> tuple[BaseException, ...]:
                assert original_release() == ()
                return (primary, cleanup)

            with monkeypatch.context() as patch:
                patch.setattr(graph, "_release_owned_read_snapshot", release_graph_failures)
                with pytest.raises(RuntimeError) as captured:
                    _invoke_surface(graph, "direct_dependents")
        elif boundary == "post-commit":
            original_finalizer = store.edge_store.finalize_transaction_commit

            def fail_commit_hook() -> None:
                raise primary

            def finalize_then_fail() -> None:
                original_finalizer()
                raise cleanup

            with monkeypatch.context() as patch:
                patch.setattr(store.edge_store, "transaction_committed", fail_commit_hook)
                patch.setattr(store.edge_store, "finalize_transaction_commit", finalize_then_fail)
                with pytest.raises(RuntimeError) as captured, store.transaction():
                    store.set_meta("prior_causal_evidence", prior_link)
        elif boundary == "constructor":
            database = store.path
            store.close()
            connection = sqlite3.connect(database, timeout=1.0, isolation_level=None)
            connection.row_factory = sqlite3.Row
            original_release = edge_module._release_initialization_transaction

            def fail_validation(
                target: sqlite3.Connection,
                backend: Any,
                tables: set[str],
            ) -> None:
                del target, backend, tables
                raise primary

            def release_constructor_failure(
                target: sqlite3.Connection,
            ) -> tuple[BaseException, ...]:
                assert original_release(target) == ()
                return (cleanup,)

            with monkeypatch.context() as patch:
                patch.setattr(edge_module, "_validate_graph_sidecar", fail_validation)
                patch.setattr(
                    edge_module,
                    "_release_initialization_transaction",
                    release_constructor_failure,
                )
                with pytest.raises(RuntimeError) as captured:
                    EdgeStore(connection)
        else:
            raise AssertionError(f"unknown cleanup boundary {boundary!r}")

        assert captured.value is primary
        _assert_unique_exception_identities(captured.value)
        assert _count_exception_identity(captured.value, prior) == 1
        assert _count_exception_identity(captured.value, cleanup) == 1
    finally:
        if connection is not None:
            connection.close()
        if not store._closed:
            store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize(
    "boundary",
    ["graph", "transaction", "exit", "post-commit", "constructor"],
)
def test_chained_boundaries_exclude_members_owned_by_primary_group(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    boundary: str,
) -> None:
    store = _seed_store(tmp_path, prefer_rtree, f"primary-graph-{boundary}")
    shared = ValueError(f"{boundary} shared primary member")
    primary_only = RuntimeError(f"{boundary} primary-only member")
    cleanup_only = OSError(f"{boundary} cleanup-only member")
    primary = BaseExceptionGroup(f"{boundary} primary", (shared, primary_only))
    cleanup = BaseExceptionGroup(f"{boundary} cleanup", (shared, cleanup_only))
    connection: sqlite3.Connection | None = None
    try:
        if boundary == "graph":
            graph = store.dependency_graph

            def fail_graph_body() -> None:
                raise primary

            def fail_graph_rollback() -> None:
                raise cleanup

            with monkeypatch.context() as patch:
                patch.setattr(graph, "_require_clean_spatial", fail_graph_body)
                patch.setattr(graph, "_rollback_owned_read_snapshot", fail_graph_rollback)
                with pytest.raises(BaseExceptionGroup) as captured:
                    _invoke_surface(graph, "direct_dependents")
        elif boundary == "transaction":

            def fail_store_rollback() -> None:
                raise cleanup

            with monkeypatch.context() as patch:
                patch.setattr(store, "_rollback_after_failed_transaction", fail_store_rollback)
                with pytest.raises(BaseExceptionGroup) as captured, store.transaction():
                    raise primary
        elif boundary == "exit":

            def fail_close() -> None:
                raise cleanup

            with monkeypatch.context() as patch:
                patch.setattr(store, "close", fail_close)
                with pytest.raises(BaseExceptionGroup) as captured, store:
                    raise primary
        elif boundary == "post-commit":
            original_finalizer = store.edge_store.finalize_transaction_commit

            def fail_commit_hook() -> None:
                raise primary

            def finalize_then_fail() -> None:
                original_finalizer()
                raise cleanup

            with monkeypatch.context() as patch:
                patch.setattr(store.edge_store, "transaction_committed", fail_commit_hook)
                patch.setattr(
                    store.edge_store,
                    "finalize_transaction_commit",
                    finalize_then_fail,
                )
                with pytest.raises(BaseExceptionGroup) as captured, store.transaction():
                    store.set_meta("primary_graph_boundary", boundary)
        elif boundary == "constructor":
            database = store.path
            store.close()
            connection = sqlite3.connect(database, timeout=1.0, isolation_level=None)
            connection.row_factory = sqlite3.Row
            original_release = edge_module._release_initialization_transaction

            def fail_validation(
                target: sqlite3.Connection,
                backend: Any,
                tables: set[str],
            ) -> None:
                del target, backend, tables
                raise primary

            def release_with_cleanup(
                target: sqlite3.Connection,
            ) -> tuple[BaseException, ...]:
                assert original_release(target) == ()
                return (cleanup,)

            with monkeypatch.context() as patch:
                patch.setattr(edge_module, "_validate_graph_sidecar", fail_validation)
                patch.setattr(
                    edge_module,
                    "_release_initialization_transaction",
                    release_with_cleanup,
                )
                with pytest.raises(BaseExceptionGroup) as captured:
                    EdgeStore(connection)
        else:
            raise AssertionError(f"unknown chained boundary {boundary!r}")

        assert captured.value is primary
        assert captured.value.__cause__ is not cleanup
        assert isinstance(captured.value.__cause__, BaseExceptionGroup)
        assert captured.value.__cause__.exceptions == (cleanup_only,)
        _assert_unique_exception_identities(captured.value)
        assert _count_exception_identity(captured.value, shared) == 1
        assert _count_exception_identity(captured.value, cleanup_only) == 1
    finally:
        if connection is not None:
            connection.close()
        if not store._closed:
            store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("boundary", ["graph", "close", "commit-hook", "constructor"])
def test_standalone_boundaries_install_rebuilt_identity_unique_groups(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    boundary: str,
) -> None:
    store = _seed_store(tmp_path, prefer_rtree, f"standalone-group-{boundary}")
    shared = ValueError(f"{boundary} shared member")
    distinct = RuntimeError(f"{boundary} distinct member")
    external = OSError(f"{boundary} external cause")
    distinct.__cause__ = external
    external.__context__ = distinct
    nested = BaseExceptionGroup(f"{boundary} nested", (shared, distinct))
    nested.add_note(f"{boundary} nested note")
    root = BaseExceptionGroup(f"{boundary} root", (shared, nested, shared, nested))
    root.add_note(f"{boundary} root note")
    root.__suppress_context__ = False

    captured_error: BaseException
    connection: sqlite3.Connection | None = None
    try:
        if boundary == "graph":
            graph = store.dependency_graph
            original_release = graph._release_owned_read_snapshot

            def release_graph_with_group() -> tuple[BaseException, ...]:
                assert original_release() == ()
                return (root,)

            with monkeypatch.context() as patch:
                patch.setattr(
                    graph,
                    "_release_owned_read_snapshot",
                    release_graph_with_group,
                )
                with pytest.raises(BaseExceptionGroup) as captured:
                    _invoke_surface(graph, "direct_dependents")
            captured_error = captured.value
        elif boundary == "close":

            def close_then_raise_group() -> None:
                sqlite3.Connection.close(store._connection)
                raise root

            with monkeypatch.context() as patch:
                patch.setattr(store, "_close_connection", close_then_raise_group)
                with pytest.raises(BaseExceptionGroup) as captured:
                    store.close()
            captured_error = captured.value
        elif boundary == "commit-hook":
            original_finalizer = store.edge_store.finalize_transaction_commit

            def finalize_then_raise_group() -> None:
                original_finalizer()
                raise root

            with monkeypatch.context() as patch:
                patch.setattr(
                    store.edge_store,
                    "transaction_committed",
                    finalize_then_raise_group,
                )
                with pytest.raises(BaseExceptionGroup) as captured, store.transaction():
                    store.set_meta("standalone_group_boundary", boundary)
            captured_error = captured.value
        elif boundary == "constructor":
            database = store.path
            store.close()
            connection = sqlite3.connect(database, timeout=1.0, isolation_level=None)
            connection.row_factory = sqlite3.Row
            original_release = edge_module._release_initialization_transaction

            def release_constructor_with_group(
                target: sqlite3.Connection,
            ) -> tuple[BaseException, ...]:
                assert original_release(target) == ()
                return (root,)

            with monkeypatch.context() as patch:
                patch.setattr(
                    edge_module,
                    "_release_initialization_transaction",
                    release_constructor_with_group,
                )
                with pytest.raises(BaseExceptionGroup) as captured:
                    EdgeStore(connection)
            captured_error = captured.value
        else:
            raise AssertionError(f"unknown standalone boundary {boundary!r}")

        assert captured_error is not root
        assert isinstance(captured_error, BaseExceptionGroup)
        assert captured_error.message == f"{boundary} root"
        assert captured_error.__notes__ == [f"{boundary} root note"]
        assert captured_error.__suppress_context__ is False
        assert len(captured_error.exceptions) == 2
        assert captured_error.exceptions[0] is shared
        normalized_nested = captured_error.exceptions[1]
        assert isinstance(normalized_nested, BaseExceptionGroup)
        assert normalized_nested.message == f"{boundary} nested"
        assert normalized_nested.__notes__ == [f"{boundary} nested note"]
        assert normalized_nested.exceptions == (distinct,)
        assert distinct.__cause__ is external
        assert external.__context__ is None
        _assert_unique_exception_identities(captured_error)
    finally:
        if connection is not None:
            connection.close()
        if not store._closed:
            store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("failure_timing", ["before", "after"])
def test_constructor_acquisition_cleanup_sanitizes_nested_primary_and_cycles(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
    failure_timing: str,
) -> None:
    store = _seed_store(
        tmp_path,
        prefer_rtree,
        f"constructor-acquisition-cleanup-{failure_timing}",
    )
    database = store.path
    store.close()
    connection = sqlite3.connect(
        database,
        timeout=1.0,
        isolation_level=None,
        factory=_BeginImmediateFailureConnection,
    )
    connection.row_factory = sqlite3.Row

    primary = sqlite3.DatabaseError(f"constructor acquisition failed {failure_timing}")
    prior = LookupError("constructor acquisition prior cause")
    primary.__cause__ = prior
    cycle_one = OSError("constructor acquisition cleanup cycle one")
    cycle_two = ValueError("constructor acquisition cleanup cycle two")
    cycle_one.__cause__ = cycle_two
    cycle_two.__context__ = cycle_one
    nested = BaseExceptionGroup("constructor acquisition nested", (primary, cycle_one))
    cleanup = BaseExceptionGroup(
        "constructor acquisition cleanup",
        (nested, nested, cycle_two),
    )
    cleanup.add_note("constructor acquisition cleanup note")
    cleanup.__suppress_context__ = False

    original_release = edge_module._release_initialization_transaction

    def release_with_nested_cycle(
        connection: sqlite3.Connection,
    ) -> tuple[BaseException, ...]:
        assert original_release(connection) == ()
        return (cleanup,)

    connection.begin_failure_timing = failure_timing
    connection.begin_failure_marker = primary
    try:
        with monkeypatch.context() as patch:
            patch.setattr(
                edge_module,
                "_release_initialization_transaction",
                release_with_nested_cycle,
            )
            with pytest.raises(sqlite3.DatabaseError) as captured:
                EdgeStore(connection)

        assert captured.value is primary
        composed = captured.value.__cause__
        assert isinstance(composed, BaseExceptionGroup)
        assert composed.message == "EdgeStore initialization causal evidence and cleanup failure"
        assert composed.exceptions[0] is prior
        normalized = composed.exceptions[1]
        assert isinstance(normalized, BaseExceptionGroup)
        assert normalized.message == "constructor acquisition cleanup"
        assert normalized.__notes__ == ["constructor acquisition cleanup note"]
        assert normalized.__suppress_context__ is False
        assert not _exception_graph_contains(normalized, primary)
        _assert_unique_exception_identities(captured.value)
        assert _count_exception_identity(captured.value, prior) == 1
        assert _count_exception_identity(captured.value, cycle_one) == 1
        assert _count_exception_identity(captured.value, cycle_two) == 1
        _assert_connection_is_inactive_or_closed(connection)
        _assert_writer_can_begin(database)
    finally:
        with suppress(sqlite3.ProgrammingError):
            connection.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_constructor_cleanup_sanitizes_nested_primary_duplicates_and_member_cycles(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    prefer_rtree: bool,
) -> None:
    store = _seed_store(tmp_path, prefer_rtree, "constructor-cleanup-sanitizer")
    database = store.path
    store.close()
    connection = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    connection.row_factory = sqlite3.Row

    primary = RuntimeError("constructor validation primary")
    distinct = LookupError("distinct constructor cleanup")
    cycle_one = OSError("constructor cleanup cycle one")
    cycle_two = ValueError("constructor cleanup cycle two")
    cycle_one.__cause__ = cycle_two
    cycle_two.__context__ = cycle_one
    inner = BaseExceptionGroup("constructor inner cleanup", (primary, cycle_one, distinct))
    inner.add_note("constructor inner note")
    inner.__suppress_context__ = True
    outer = BaseExceptionGroup("constructor outer cleanup", (inner, inner, cycle_two))
    outer.add_note("constructor outer note")
    outer.__suppress_context__ = True

    original_release = edge_module._release_initialization_transaction

    def fail_validation(
        connection: sqlite3.Connection,
        backend: Any,
        tables: set[str],
    ) -> None:
        del connection, backend, tables
        raise primary

    def release_with_nested_cycle(
        connection: sqlite3.Connection,
    ) -> tuple[BaseException, ...]:
        assert original_release(connection) == ()
        return (outer,)

    try:
        with monkeypatch.context() as patch:
            patch.setattr(edge_module, "_validate_graph_sidecar", fail_validation)
            patch.setattr(
                edge_module,
                "_release_initialization_transaction",
                release_with_nested_cycle,
            )
            with pytest.raises(RuntimeError) as captured:
                EdgeStore(connection)

        assert captured.value is primary
        assert not connection.in_transaction
        cleanup = captured.value.__cause__
        assert isinstance(cleanup, BaseExceptionGroup)
        assert cleanup.message == "constructor outer cleanup"
        assert cleanup.__notes__ == ["constructor outer note"]
        assert cleanup.__suppress_context__ is True
        assert len(cleanup.exceptions) == 2
        normalized_inner = cleanup.exceptions[0]
        assert isinstance(normalized_inner, BaseExceptionGroup)
        assert normalized_inner.message == "constructor inner cleanup"
        assert normalized_inner.__notes__ == ["constructor inner note"]
        assert normalized_inner.__suppress_context__ is True
        assert normalized_inner.exceptions == (cycle_one, distinct)
        assert cleanup.exceptions[1] is cycle_two
        assert not _exception_graph_contains(cleanup, primary)
        _assert_unique_exception_identities(captured.value)
        assert _count_exception_identity(captured.value, cycle_one) == 1
        assert _count_exception_identity(captured.value, cycle_two) == 1
        assert _count_exception_identity(captured.value, distinct) == 1
        _assert_writer_can_begin(database)
    finally:
        connection.close()


def _seed_store(
    tmp_path: Path,
    prefer_rtree: bool,
    suffix: str,
) -> IndexStore:
    store = IndexStore(
        tmp_path / f"p4-mechanics-{prefer_rtree}-{suffix}.xlsp.db",
        prefer_rtree=prefer_rtree,
    )
    try:
        store.replace_sheet_catalog(
            (
                _sheet(1, "Inputs", 0),
                _sheet(2, "Calc", 1),
                _sheet(3, "Summary", 2),
            )
        )
        store.connection.executemany(
            "INSERT INTO fblocks VALUES (?, ?, ?, '=RC', ?, ?, ?, ?, 0, 0)",
            (
                (11, 2, 0, 1, 2, 2, 2),
                (13, 3, 0, 1, 1, 3, 3),
            ),
        )
        store.connection.executemany(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (?, 'fblock', ?, ?, ?, ?, ?, ?, ?, 'ref')
            """,
            (
                (1, 11, 2, 1, 1, 1_048_576, 1, 1),
                (2, 13, 3, 2, 1, 2, 2, 2),
                # A valid duplicate semantic hop. Edge 1 remains the representative;
                # mutating only edge 3 probes whether bounded queries can hide a split.
                (3, 11, 2, 1, 1, 1_048_576, 1, 1),
            ),
        )
        store.rebuild_graph_spatial_index()
    except BaseException:
        store.close()
        raise
    return store


def _sheet(sheet_id: int, name: str, order: int) -> SheetDescriptor:
    return SheetDescriptor(
        order=order,
        name=name,
        sheet_id=sheet_id,
        rel_id=f"rId{sheet_id}",
        xml_part=f"xl/worksheets/sheet{sheet_id}.xml",
        kind="worksheet",
    )


def _invoke_surface(graph: DependencyGraph, surface: str) -> object:
    if surface == "direct_precedents":
        return graph.direct_precedents(GraphArea(2, "Calc", Rect(1, 2, 2, 2)))
    if surface == "direct_dependents":
        return graph.direct_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
    if surface == "trace_precedents":
        return graph.trace_precedents(GraphArea(3, "Summary", Rect(1, 1, 3, 3)))
    if surface == "trace_dependents":
        return graph.trace_dependents(GraphArea(1, "Inputs", Rect(1, 1, 1, 1)))
    if surface == "trace_path":
        return graph.trace_path(
            GraphArea(1, "Inputs", Rect(1, 1, 1, 1)),
            GraphArea(3, "Summary", Rect(1, 1, 3, 3)),
        )
    raise AssertionError(f"unknown graph surface {surface!r}")


def _split_duplicate_and_restore_persisted_seals(store: IndexStore, *, mode: str) -> None:
    state = _graph_state(store.connection)
    catalog = _catalog(store.connection)
    _mutate_duplicate(store.connection, mode, via="coherent-split")
    store.connection.executemany(
        "INSERT OR REPLACE INTO graph_rank_keys(direction, rank, key_text) VALUES (?, ?, ?)",
        catalog,
    )
    _restore_graph_state(store.connection, state)


def _mutate_duplicate(connection: SQLiteConnectionLike, mode: str, *, via: str) -> None:
    sql = "UPDATE edges SET via = ? WHERE id = ?"
    if mode == "execute":
        connection.execute(sql, (via, 3))
    elif mode == "cursor":
        cursor = connection.cursor()
        try:
            cursor.execute(sql, (via, 3))
        finally:
            cursor.close()
    elif mode == "executemany":
        connection.executemany(sql, ((via, 3),))
    else:
        raise AssertionError(f"unknown execution mode {mode!r}")


def _graph_state(connection: SQLiteConnectionLike) -> tuple[int, ...]:
    row = connection.execute(
        """
        SELECT singleton, dirty, dependent_rank_max, precedent_rank_max,
               revision, mutation_epoch, clean_epoch
        FROM graph_spatial_state
        """
    ).fetchone()
    assert row is not None
    return cast(tuple[int, ...], tuple(row))


def _catalog(connection: SQLiteConnectionLike) -> tuple[tuple[object, ...], ...]:
    return tuple(
        tuple(row)
        for row in connection.execute(
            "SELECT direction, rank, key_text FROM graph_rank_keys ORDER BY direction, rank"
        )
    )


def _edge_rows(connection: SQLiteConnectionLike) -> tuple[tuple[object, ...], ...]:
    return tuple(
        tuple(row)
        for row in connection.execute(
            "SELECT id, src_kind, src_id, src_sheet_id, dst_sheet_id, "
            "dst_row_min, dst_row_max, dst_col_min, dst_col_max, via, "
            "dependent_rank, precedent_rank FROM edges ORDER BY id"
        )
    )


def _graph_write_epoch(connection: SQLiteConnectionLike) -> int:
    epoch = getattr(connection, "_graph_write_epoch", None)
    assert type(epoch) is int
    return epoch


def _restore_graph_state(connection: SQLiteConnectionLike, state: tuple[int, ...]) -> None:
    connection.execute(
        """
        UPDATE graph_spatial_state
        SET singleton = ?, dirty = ?, dependent_rank_max = ?, precedent_rank_max = ?,
            revision = ?, mutation_epoch = ?, clean_epoch = ?
        """,
        state,
    )


class _BeginImmediateFailureConnection(sqlite3.Connection):
    begin_failure_timing: str | None = None
    begin_failure_marker: sqlite3.DatabaseError | None = None

    def execute(self, sql: str, parameters: Any = (), /) -> sqlite3.Cursor:
        if self.begin_failure_timing is not None and sql.strip().upper() == "BEGIN IMMEDIATE":
            timing = self.begin_failure_timing
            self.begin_failure_timing = None
            if timing == "after":
                super().execute(sql, parameters)
            marker = self.begin_failure_marker
            assert marker is not None
            raise marker
        return super().execute(sql, parameters)


def _assert_connection_is_inactive_or_closed(connection: sqlite3.Connection) -> None:
    descriptor = cast(Any, sqlite3.Connection.in_transaction)
    try:
        active = bool(descriptor.__get__(connection, sqlite3.Connection))
    except sqlite3.ProgrammingError as error:
        assert "closed" in str(error).lower()
    else:
        assert not active


def _assert_writer_can_begin(database: Path) -> None:
    writer = sqlite3.connect(database, timeout=1.0, isolation_level=None)
    try:
        writer.execute("PRAGMA busy_timeout = 1000")
        writer.execute("BEGIN IMMEDIATE")
        writer.rollback()
    finally:
        writer.close()


def _assert_unique_exception_identities(root: BaseException) -> None:
    seen: set[int] = set()
    pending = [root]
    while pending:
        error = pending.pop()
        assert id(error) not in seen, f"duplicate or cyclic exception identity: {error!r}"
        seen.add(id(error))
        pending.extend(_exception_links(error))


def _count_exception_identity(root: BaseException, target: BaseException) -> int:
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
        pending.extend(_exception_links(error))
    return count


def _exception_graph_contains(root: BaseException, target: BaseException) -> bool:
    expanded: set[int] = set()
    pending = [root]
    while pending:
        error = pending.pop()
        if error is target:
            return True
        if id(error) in expanded:
            continue
        expanded.add(id(error))
        pending.extend(_exception_links(error))
    return False


def _exception_links(error: BaseException) -> list[BaseException]:
    links: list[BaseException] = []
    if error.__cause__ is not None:
        links.append(error.__cause__)
    if error.__context__ is not None:
        links.append(error.__context__)
    if isinstance(error, BaseExceptionGroup):
        links.extend(cast(tuple[BaseException, ...], error.exceptions))
    return links
