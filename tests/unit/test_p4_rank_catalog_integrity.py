from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import cast

import pytest

from excel_lsp.core.index.edges import (
    EdgeSchemaError,
    EdgeStore,
    SQLiteConnectionLike,
    canonical_rank_key_text,
    canonical_ranked_edges,
)
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import SheetDescriptor

_CATALOG_DDL = """
CREATE TABLE graph_rank_keys (
    direction TEXT NOT NULL CHECK (direction IN ('dependents', 'precedents')),
    rank INTEGER NOT NULL CHECK (rank > 0),
    key_text TEXT NOT NULL,
    PRIMARY KEY (direction, rank)
) WITHOUT ROWID
"""

_FROZEN_BASE_TRIGGER_MANIFEST = (
    ("edges_graph_spatial_dirty_insert", "edges", "insert"),
    ("edges_graph_spatial_dirty_update", "edges", "update"),
    ("edges_graph_spatial_dirty_delete", "edges", "delete"),
    ("fblocks_graph_spatial_dirty_insert", "fblocks", "insert"),
    ("fblocks_graph_spatial_dirty_update", "fblocks", "update"),
    ("fblocks_graph_spatial_dirty_delete", "fblocks", "delete"),
    ("sheets_graph_spatial_dirty_insert", "sheets", "insert"),
    ("sheets_graph_spatial_dirty_update", "sheets", "update"),
    ("sheets_graph_spatial_dirty_delete", "sheets", "delete"),
    (
        "graph_rank_keys_graph_spatial_dirty_insert",
        "graph_rank_keys",
        "insert",
    ),
    (
        "graph_rank_keys_graph_spatial_dirty_update",
        "graph_rank_keys",
        "update",
    ),
    (
        "graph_rank_keys_graph_spatial_dirty_delete",
        "graph_rank_keys",
        "delete",
    ),
)

_FROZEN_MIRROR_TRIGGER_MANIFEST = {
    "rtree": (
        ("edge_rtree_rowid_graph_dirty_insert", "edge_rtree_rowid", "insert"),
        ("edge_rtree_rowid_graph_dirty_update", "edge_rtree_rowid", "update"),
        ("edge_rtree_rowid_graph_dirty_delete", "edge_rtree_rowid", "delete"),
        (
            "edge_source_rtree_rowid_graph_dirty_insert",
            "edge_source_rtree_rowid",
            "insert",
        ),
        (
            "edge_source_rtree_rowid_graph_dirty_update",
            "edge_source_rtree_rowid",
            "update",
        ),
        (
            "edge_source_rtree_rowid_graph_dirty_delete",
            "edge_source_rtree_rowid",
            "delete",
        ),
    ),
    "interval": (
        ("edge_intervals_graph_dirty_insert", "edge_intervals", "insert"),
        ("edge_intervals_graph_dirty_update", "edge_intervals", "update"),
        ("edge_intervals_graph_dirty_delete", "edge_intervals", "delete"),
        (
            "edge_source_intervals_graph_dirty_insert",
            "edge_source_intervals",
            "insert",
        ),
        (
            "edge_source_intervals_graph_dirty_update",
            "edge_source_intervals",
            "update",
        ),
        (
            "edge_source_intervals_graph_dirty_delete",
            "edge_source_intervals",
            "delete",
        ),
    ),
}

_CATALOG_DDL_MALFORMATIONS = (
    "rowid",
    "extra-column",
    "wrong-rank-constraint",
    "wrong-direction-constraint",
    "reversed-primary-key",
    "single-column-primary-key",
    "nullable-direction",
    "nullable-rank",
    "nullable-key-text",
)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_rank_catalog_has_exact_frozen_physical_identity(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    with _seed_store(tmp_path, prefer_rtree, "exact-ddl", generation=10) as store:
        row = store.connection.execute(
            "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'graph_rank_keys'"
        ).fetchone()
        assert row is not None
        assert _canonical_sql(str(row[0])) == _canonical_sql(_CATALOG_DDL)
        assert tuple(
            (str(item[1]), str(item[2]).casefold(), int(item[3]), int(item[5]))
            for item in store.connection.execute("PRAGMA table_info(graph_rank_keys)")
        ) == (
            ("direction", "text", 1, 1),
            ("rank", "integer", 1, 2),
            ("key_text", "text", 1, 0),
        )
        assert _catalog_rows(store.connection) == _canonical_catalog_rows(store.connection)
        store.edge_store.require_clean()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("malformation", _CATALOG_DDL_MALFORMATIONS)
def test_malformed_rank_catalog_ddl_is_rejected_and_rebuilt_monotonically(
    tmp_path: Path,
    prefer_rtree: bool,
    malformation: str,
) -> None:
    database = tmp_path / f"catalog-ddl-{prefer_rtree}-{malformation}.xlsp.db"
    with _seed_store_at(database, prefer_rtree, generation=20) as store:
        rows = _catalog_rows(store.connection)
        before_generation = store.generation

    connection = sqlite3.connect(database, isolation_level=None)
    try:
        connection.execute("DROP TABLE graph_rank_keys")
        ddl = _malformed_catalog_ddl(malformation)
        connection.execute(ddl)
        connection.executemany(
            "INSERT INTO graph_rank_keys(direction, rank, key_text) VALUES (?, ?, ?)",
            rows,
        )
        with pytest.raises(EdgeSchemaError, match=r"physical identity|malformed columns"):
            EdgeStore.ensure_schema(connection, prefer_rtree=prefer_rtree)
        with pytest.raises(EdgeSchemaError, match=r"physical identity|malformed columns"):
            EdgeStore(connection)
    finally:
        connection.close()

    with IndexStore(database, prefer_rtree=prefer_rtree) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == before_generation + 1
        assert _catalog_rows(rebuilt.connection) == _canonical_catalog_rows(rebuilt.connection)
        row = rebuilt.connection.execute(
            "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'graph_rank_keys'"
        ).fetchone()
        assert row is not None
        assert _canonical_sql(str(row[0])) == _canonical_sql(_CATALOG_DDL)
        assert rebuilt.edge_store.backend == _expected_backend(prefer_rtree)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("damage", ["missing", "wrong", "extra"])
@pytest.mark.parametrize("direction", ["dependents", "precedents"])
def test_rank_catalog_content_damage_survives_state_restoration_only_to_be_rebuilt(
    tmp_path: Path,
    prefer_rtree: bool,
    damage: str,
    direction: str,
) -> None:
    database = tmp_path / f"catalog-content-{prefer_rtree}-{direction}-{damage}.xlsp.db"
    with _seed_store_at(database, prefer_rtree, generation=30) as store:
        sealed_state = _graph_state(store.connection)
        before_generation = store.generation

    connection = sqlite3.connect(database, isolation_level=None)
    try:
        if damage == "missing":
            connection.execute(
                "DELETE FROM graph_rank_keys WHERE direction = ? AND rank = 1",
                (direction,),
            )
        elif damage == "wrong":
            connection.execute(
                "UPDATE graph_rank_keys SET key_text = 'forged' WHERE direction = ? AND rank = 1",
                (direction,),
            )
        else:
            connection.execute(
                "INSERT INTO graph_rank_keys(direction, rank, key_text) "
                "VALUES (?, 99, 'forged-extra')",
                (direction,),
            )
        _restore_graph_state(connection, sealed_state)
        with pytest.raises(EdgeSchemaError, match="rank-key catalog"):
            EdgeStore.ensure_schema(connection, prefer_rtree=prefer_rtree)
        with pytest.raises(EdgeSchemaError, match="rank-key catalog"):
            EdgeStore(connection)
    finally:
        connection.close()

    with IndexStore(database, prefer_rtree=prefer_rtree) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == before_generation + 1
        assert _catalog_rows(rebuilt.connection) == _canonical_catalog_rows(rebuilt.connection)
        assert rebuilt.edge_store.backend == _expected_backend(prefer_rtree)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
def test_every_rank_invalidation_trigger_is_exact_and_killable(
    tmp_path: Path,
    prefer_rtree: bool,
) -> None:
    database = tmp_path / f"catalog-triggers-{prefer_rtree}.xlsp.db"
    with _seed_store_at(database, prefer_rtree, generation=40) as store:
        before_generation = store.generation

    connection = sqlite3.connect(database, isolation_level=None)
    try:
        backend = _expected_backend(prefer_rtree)
        manifest = (
            *_FROZEN_BASE_TRIGGER_MANIFEST,
            *_FROZEN_MIRROR_TRIGGER_MANIFEST[backend],
        )
        expected_by_name = {
            name: (table, _frozen_trigger_sql(name, table, operation))
            for name, table, operation in manifest
        }
        triggers = tuple(
            (str(row[0]), str(row[1]), str(row[2]))
            for row in connection.execute(
                """
                SELECT name, tbl_name, sql FROM sqlite_master
                WHERE type = 'trigger'
                  AND (name LIKE '%graph_spatial_dirty_%' OR name LIKE '%graph_dirty_%')
                ORDER BY name
                """
            )
        )
        assert EdgeStore.ensure_schema(connection, prefer_rtree=prefer_rtree) == backend
        assert tuple((name, table) for name, table, _sql in triggers) == tuple(
            sorted((name, table) for name, table, _operation in manifest)
        )
        for name, table, sql in triggers:
            expected_table, expected_sql = expected_by_name[name]
            assert table == expected_table
            assert _canonical_sql(sql) == _canonical_sql(expected_sql)

        for name, table, operation in manifest:
            sql = _frozen_trigger_sql(name, table, operation)
            connection.execute(f'DROP TRIGGER "{name}"')
            with pytest.raises(EdgeSchemaError, match=r"trigger.*missing"):
                EdgeStore(connection)
            connection.execute(sql)
            EdgeStore(connection).require_clean()

            connection.execute(f'DROP TRIGGER "{name}"')
            mutated = sql.replace("BEGIN", "/* killed exact body */ BEGIN", 1)
            assert mutated != sql
            connection.execute(mutated)
            with pytest.raises(EdgeSchemaError, match=r"trigger.*malformed"):
                EdgeStore(connection)
            connection.execute(f'DROP TRIGGER "{name}"')
            connection.execute(sql)
            EdgeStore(connection).require_clean()

        connection.execute(f'DROP TRIGGER "{manifest[0][0]}"')
    finally:
        connection.close()

    with IndexStore(database, prefer_rtree=prefer_rtree) as rebuilt:
        assert rebuilt.schema_rebuilt is True
        assert rebuilt.generation == before_generation + 1
        assert _catalog_rows(rebuilt.connection) == _canonical_catalog_rows(rebuilt.connection)
        trigger_count = int(
            rebuilt.connection.execute(
                """
                SELECT COUNT(*) FROM sqlite_master
                WHERE type = 'trigger'
                  AND (name LIKE '%graph_spatial_dirty_%' OR name LIKE '%graph_dirty_%')
                """
            ).fetchone()[0]
        )
        assert trigger_count == 18
        assert rebuilt.edge_store.backend == _expected_backend(prefer_rtree)


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("operation", ["insert", "update", "delete"])
@pytest.mark.parametrize("table", ["edges", "fblocks", "sheets", "graph_rank_keys"])
def test_base_catalog_trigger_families_functionally_invalidate_clean_graphs(
    tmp_path: Path,
    prefer_rtree: bool,
    operation: str,
    table: str,
) -> None:
    store = _seed_store(tmp_path, prefer_rtree, f"base-{table}-{operation}", generation=50)
    try:
        before_catalog = _catalog_rows(store.connection)
        assert before_catalog
        assert _graph_state(store.connection)[1] == 0
        store.connection.execute("PRAGMA foreign_keys = OFF")
        if table == "edges":
            if operation == "insert":
                store.connection.execute(
                    """
                    INSERT INTO edges
                    SELECT 2, src_kind, src_id, src_sheet_id, dst_sheet_id,
                           dst_row_min, dst_row_max, dst_col_min, dst_col_max,
                           via, dependent_rank, precedent_rank
                    FROM edges WHERE id = 1
                    """
                )
            elif operation == "update":
                store.connection.execute("UPDATE edges SET via = via WHERE id = 1")
            else:
                store.connection.execute("DELETE FROM edges WHERE id = 1")
        elif table == "fblocks":
            if operation == "insert":
                store.connection.execute(
                    "INSERT INTO fblocks VALUES (2, 1, 1, '=RC', 2, 2, 1, 1, 0, 0)"
                )
            elif operation == "update":
                store.connection.execute("UPDATE fblocks SET r1c1 = r1c1 WHERE id = 1")
            else:
                store.connection.execute("DELETE FROM fblocks WHERE id = 1")
        elif table == "sheets":
            if operation == "insert":
                store.connection.execute(
                    "INSERT INTO sheets VALUES (3, 'Extra', 'xl/worksheets/sheet3.xml', "
                    "'hash', 'worksheet', 'visible', 1, 1)"
                )
            elif operation == "update":
                store.connection.execute("UPDATE sheets SET name = name WHERE id = 1")
            else:
                store.connection.execute("DELETE FROM sheets WHERE id = 2")
        elif operation == "insert":
            store.connection.execute(
                "INSERT INTO graph_rank_keys VALUES ('dependents', 99, 'extra')"
            )
        elif operation == "update":
            store.connection.execute(
                "UPDATE graph_rank_keys SET key_text = key_text || '-mutated' "
                "WHERE direction = 'dependents' AND rank = 1"
            )
        else:
            store.connection.execute(
                "DELETE FROM graph_rank_keys WHERE direction = 'dependents' AND rank = 1"
            )

        state = _graph_state(store.connection)
        assert state[1] == 1
        assert int(state[5]) > int(state[6])
        actual_catalog = _catalog_rows(store.connection)
        if table != "graph_rank_keys":
            assert actual_catalog == ()
        elif operation == "insert":
            assert actual_catalog == tuple(sorted((*before_catalog, ("dependents", 99, "extra"))))
        elif operation == "update":
            assert actual_catalog == tuple(
                (
                    direction,
                    rank,
                    f"{key_text}-mutated" if (direction, rank) == ("dependents", 1) else key_text,
                )
                for direction, rank, key_text in before_catalog
            )
        else:
            assert actual_catalog == tuple(
                row for row in before_catalog if row[:2] != ("dependents", 1)
            )
    finally:
        store.close()


@pytest.mark.parametrize("prefer_rtree", [True, False], ids=["rtree", "interval"])
@pytest.mark.parametrize("operation", ["insert", "update", "delete"])
@pytest.mark.parametrize("direction", ["dependents", "precedents"])
def test_active_mirror_trigger_families_functionally_invalidate_rank_identity(
    tmp_path: Path,
    prefer_rtree: bool,
    operation: str,
    direction: str,
) -> None:
    store = _seed_store(
        tmp_path,
        prefer_rtree,
        f"mirror-{direction}-{operation}",
        generation=60,
    )
    try:
        before_catalog = _catalog_rows(store.connection)
        table = (
            store.edge_store.table_name
            if direction == "dependents"
            else store.edge_store.source_table_name
        )
        assert (
            store.connection.execute(
                "SELECT 1 FROM graph_rank_keys WHERE direction = ? AND rank = 1",
                (direction,),
            ).fetchone()
            is not None
        )
        columns = tuple(
            str(row[1]) for row in store.connection.execute(f"PRAGMA table_info({table})")
        )
        assert columns[0] == "edge_id"
        if operation == "insert":
            projection = ", ".join(("99", *columns[1:]))
            store.connection.execute(
                f"INSERT INTO {table} SELECT {projection} FROM {table} WHERE edge_id = 1"
            )
        elif operation == "update":
            store.connection.execute(f"UPDATE {table} SET row_min = row_min WHERE edge_id = 1")
        else:
            store.connection.execute(f"DELETE FROM {table} WHERE edge_id = 1")

        state = _graph_state(store.connection)
        assert state[1] == 1
        assert int(state[5]) > int(state[6])
        assert _catalog_rows(store.connection) == tuple(
            row for row in before_catalog if row[0] != direction
        )
    finally:
        store.close()


def _seed_store(tmp_path: Path, prefer_rtree: bool, suffix: str, *, generation: int) -> IndexStore:
    return _seed_store_at(tmp_path / f"rank-catalog-{suffix}.xlsp.db", prefer_rtree, generation)


def _seed_store_at(database: Path, prefer_rtree: bool, generation: int) -> IndexStore:
    store = IndexStore(database, prefer_rtree=prefer_rtree)
    try:
        assert store.edge_store.backend == _expected_backend(prefer_rtree)
        store.set_meta("generation", generation)
        store.replace_sheet_catalog((_sheet(1, "Source", 0), _sheet(2, "Destination", 1)))
        store.connection.execute("INSERT INTO fblocks VALUES (1, 1, 0, '=RC', 1, 1, 1, 1, 0, 0)")
        store.connection.execute(
            """
            INSERT INTO edges(
                id, src_kind, src_id, src_sheet_id, dst_sheet_id,
                dst_row_min, dst_row_max, dst_col_min, dst_col_max, via
            ) VALUES (1, 'fblock', 1, 1, 2, 2, 2, 2, 2, 'ref')
            """
        )
        store.rebuild_graph_spatial_index()
    except BaseException:
        store.close()
        raise
    return store


def _expected_backend(prefer_rtree: bool) -> str:
    return "rtree" if prefer_rtree else "interval"


def _malformed_catalog_ddl(malformation: str) -> str:
    if malformation == "rowid":
        return _CATALOG_DDL.replace(" WITHOUT ROWID", "")
    if malformation == "extra-column":
        return _CATALOG_DDL.replace(
            "key_text TEXT NOT NULL,",
            "key_text TEXT NOT NULL, extra_identity TEXT,",
        )
    if malformation == "wrong-rank-constraint":
        return _CATALOG_DDL.replace("CHECK (rank > 0)", "CHECK (rank >= 0)")
    if malformation == "wrong-direction-constraint":
        return _CATALOG_DDL.replace(
            "CHECK (direction IN ('dependents', 'precedents'))",
            "CHECK (direction <> '')",
        )
    if malformation == "reversed-primary-key":
        return _CATALOG_DDL.replace(
            "PRIMARY KEY (direction, rank)",
            "PRIMARY KEY (rank, direction)",
        )
    if malformation == "single-column-primary-key":
        return _CATALOG_DDL.replace(
            "PRIMARY KEY (direction, rank)",
            "PRIMARY KEY (direction)",
        )
    if malformation == "nullable-direction":
        return _CATALOG_DDL.replace("direction TEXT NOT NULL", "direction TEXT")
    if malformation == "nullable-rank":
        return _CATALOG_DDL.replace("rank INTEGER NOT NULL", "rank INTEGER")
    if malformation == "nullable-key-text":
        return _CATALOG_DDL.replace("key_text TEXT NOT NULL", "key_text TEXT")
    raise AssertionError(f"unknown catalog DDL malformation: {malformation}")


def _frozen_trigger_sql(name: str, table: str, operation: str) -> str:
    invalidation = _frozen_catalog_invalidation_sql(table, operation)
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


def _frozen_catalog_invalidation_sql(table: str, operation: str) -> str:
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
    if table in {"edge_intervals", "edge_source_intervals"}:
        direction = "precedents" if table.startswith("edge_source") else "dependents"
        if operation == "insert":
            predicate = "rank = NEW.rank"
        elif operation == "update":
            predicate = "rank IN (OLD.rank, NEW.rank)"
        else:
            predicate = "rank = OLD.rank"
        return f"DELETE FROM graph_rank_keys WHERE direction = '{direction}' AND {predicate};"
    if table in {"edge_rtree_rowid", "edge_source_rtree_rowid"}:
        direction = "precedents" if table.startswith("edge_source") else "dependents"
        return f"DELETE FROM graph_rank_keys WHERE direction = '{direction}';"
    if table == "graph_rank_keys":
        return ""
    raise AssertionError(f"unknown frozen trigger table: {table}")


def _sheet(sheet_id: int, name: str, order: int) -> SheetDescriptor:
    return SheetDescriptor(
        order=order,
        name=name,
        sheet_id=sheet_id,
        rel_id=f"rId{sheet_id}",
        xml_part=f"xl/worksheets/sheet{sheet_id}.xml",
        kind="worksheet",
    )


def _catalog_rows(connection: SQLiteConnectionLike) -> tuple[tuple[object, ...], ...]:
    return tuple(
        tuple(row)
        for row in connection.execute(
            "SELECT direction, rank, key_text FROM graph_rank_keys ORDER BY direction, rank"
        )
    )


def _canonical_catalog_rows(connection: SQLiteConnectionLike) -> tuple[tuple[object, ...], ...]:
    keys: dict[tuple[str, int], str] = {}
    for record in canonical_ranked_edges(connection, require_persisted_ranks=True):
        keys[("dependents", record.dependent_rank)] = canonical_rank_key_text(record.dependent_key)
        keys[("precedents", record.precedent_rank)] = canonical_rank_key_text(record.precedent_key)
    return tuple((direction, rank, key) for (direction, rank), key in sorted(keys.items()))


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


def _restore_graph_state(
    connection: SQLiteConnectionLike,
    state: tuple[int, ...],
) -> None:
    connection.execute(
        """
        UPDATE graph_spatial_state
        SET singleton = ?, dirty = ?, dependent_rank_max = ?, precedent_rank_max = ?,
            revision = ?, mutation_epoch = ?, clean_epoch = ?
        """,
        state,
    )


def _canonical_sql(sql: str) -> str:
    return "".join(sql.casefold().split()).rstrip(";")
