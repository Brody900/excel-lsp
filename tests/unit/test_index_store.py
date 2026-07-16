"""Focused SQLite store, canonical export, and spatial-backend tests."""

from __future__ import annotations

import sqlite3
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from threading import Barrier

import pytest

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.index.edges import EdgeStore
from excel_lsp.core.index.schema import SCHEMA_VERSION
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import (
    CellRecord,
    DataValidationInfo,
    Rect,
    SheetDescriptor,
    SheetParseSummary,
)


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
            store.edge_store.table_name,
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
        assert "sheet_id, row_min, row_max, col_min, col_max" in index_sql
    finally:
        fallback_connection.close()


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


def _exercise_edges(edges: EdgeStore) -> None:
    edges.insert(9, 1, Rect(1, 10, 2, 2))
    edges.insert(4, 1, Rect(20, 30, 2, 4))
    edges.insert(7, 2, Rect(1, 100, 1, 10))
    assert edges.query_point(1, 5, 2) == (9,)
    assert edges.query_point(1, 15, 2) == ()
    assert edges.query_range(1, Rect(8, 22, 1, 3)) == (4, 9)
    assert edges.query_range(2, Rect(50, 50, 5, 5)) == (7,)
    edges.delete(9)
    assert edges.query_point(1, 5, 2) == ()
