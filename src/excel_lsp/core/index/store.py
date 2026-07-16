"""Persistent per-workbook SQLite index storage."""

from __future__ import annotations

import json
import math
import sqlite3
import time
from collections.abc import Callable, Generator, Mapping, Sequence
from contextlib import contextmanager
from pathlib import Path
from types import TracebackType
from typing import Self

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.index.edges import EdgeStore
from excel_lsp.core.index.schema import (
    BASE_SCHEMA_SQL,
    CONTENT_TABLES_DELETE_ORDER,
    SCHEMA_VERSION,
)
from excel_lsp.core.models import (
    CellRecord,
    DataTableFormulaInfo,
    SheetDescriptor,
    SheetParseSummary,
    WorkbookMetadata,
)
from excel_lsp.core.values import JsonScalar, normalize_value

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


class IndexStore:
    """One SQLite index with deterministic schema and transaction helpers."""

    def __init__(self, path: str | Path, *, prefer_rtree: bool = True) -> None:
        self.path = Path(path).expanduser().resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._connection = sqlite3.connect(
            self.path,
            timeout=5.0,
            isolation_level=None,
        )
        self._connection.row_factory = sqlite3.Row
        self._closed = False
        self.schema_created = False
        self.schema_rebuilt = False
        try:
            self._connection.execute("PRAGMA busy_timeout = 5000")
            _retry_on_busy(lambda: self._connection.execute("PRAGMA journal_mode = WAL").fetchone())
            _retry_on_busy(
                lambda: self._connection.execute("PRAGMA synchronous = NORMAL").fetchone()
            )
            self._connection.execute("PRAGMA foreign_keys = ON")
            self._initialize_schema(prefer_rtree=prefer_rtree)
            self.edge_store = EdgeStore(self._connection)
        except BaseException:
            self._connection.close()
            self._closed = True
            raise

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        del exc_type, exc_value, traceback
        self.close()

    @property
    def connection(self) -> sqlite3.Connection:
        """Expose the connection for higher core layers using the frozen schema."""
        if self._closed:
            raise RuntimeError("index store is closed")
        return self._connection

    @property
    def generation(self) -> int:
        """Return the current monotonically increasing index generation."""
        raw = self.get_meta("generation", "0")
        try:
            return int(raw or "0")
        except ValueError as exc:
            raise RuntimeError("index generation is not an integer") from exc

    def close(self) -> None:
        """Rollback any abandoned transaction and close the connection."""
        if self._closed:
            return
        if self._connection.in_transaction:
            self._connection.rollback()
        self._connection.close()
        self._closed = True

    @contextmanager
    def transaction(self) -> Generator[None, None, None]:
        """Run an immediate atomic mutation, nesting within an existing one."""
        if self._connection.in_transaction:
            yield
            return
        self._connection.execute("BEGIN IMMEDIATE")
        try:
            yield
        except BaseException:
            self._connection.rollback()
            raise
        else:
            self._connection.commit()

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

    def replace_sheet(
        self,
        descriptor: SheetDescriptor,
        parse_sheet: SheetParser,
        *,
        batch_size: int = 1_000,
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

    def canonical_export(self) -> dict[str, tuple[tuple[object, ...], ...]]:
        """Return content rows in natural order with local row ids projected out.

        Lifecycle-only ``generation`` and ``indexed_at`` metadata are omitted so
        a full build and an incremental build of the same workbook are directly
        comparable for invariant I2.
        """
        queries = {
            "meta": """
                SELECT key, value FROM meta
                WHERE key NOT IN ('generation', 'indexed_at')
                ORDER BY key
            """,
            "sheets": """
                SELECT name, xml_part, part_hash, kind, visibility, max_row, max_col
                FROM sheets ORDER BY name
            """,
            "regions": """
                SELECT s.name, r.n, r.row_min, r.row_max, r.col_min, r.col_max,
                       r.header_rows, r.kind, r.list_object_name, r.confidence
                FROM regions AS r JOIN sheets AS s ON s.id = r.sheet_id
                ORDER BY s.name, r.n
            """,
            "columns": """
                SELECT s.name, r.n, c.idx, c.header, c.norm_header, c.dtype,
                       c.nonnull, c.distinct_est, fb.n
                FROM columns AS c
                JOIN regions AS r ON r.id = c.region_id
                JOIN sheets AS s ON s.id = r.sheet_id
                LEFT JOIN fblocks AS fb ON fb.id = c.formula_block_id
                ORDER BY s.name, r.n, c.idx
            """,
            "fblocks": """
                SELECT s.name, f.n, f.r1c1, f.row_min, f.row_max, f.col_min,
                       f.col_max, f.volatile, f.opaque
                FROM fblocks AS f JOIN sheets AS s ON s.id = f.sheet_id
                ORDER BY s.name, f.n
            """,
            "defined_names": """
                SELECT d.name, scope.name, d.refers_to, d.kind, d.is_builtin
                FROM defined_names AS d
                LEFT JOIN sheets AS scope ON scope.id = d.scope_sheet_id
                ORDER BY COALESCE(scope.name, ''), d.name, d.refers_to
            """,
            "name_areas": """
                SELECT d.name, scope.name, d.refers_to, area_sheet.name,
                       a.row_min, a.row_max, a.col_min, a.col_max
                FROM name_areas AS a
                JOIN defined_names AS d ON d.id = a.name_id
                LEFT JOIN sheets AS scope ON scope.id = d.scope_sheet_id
                JOIN sheets AS area_sheet ON area_sheet.id = a.sheet_id
                ORDER BY COALESCE(scope.name, ''), d.name, d.refers_to,
                         area_sheet.name, a.row_min, a.col_min, a.row_max, a.col_max
            """,
            "validations": """
                SELECT s.name, v.row_min, v.row_max, v.col_min, v.col_max,
                       v.vtype, v.operator, v.formula1, v.formula2, v.allow_blank
                FROM validations AS v JOIN sheets AS s ON s.id = v.sheet_id
                ORDER BY s.name, v.row_min, v.col_min, v.row_max, v.col_max,
                         v.vtype, v.operator, v.formula1, v.formula2, v.allow_blank
            """,
            "edges": """
                SELECT e.src_kind, src_sheet.name,
                       CASE WHEN e.src_kind = 'fblock' THEN src_block.n ELSE e.src_id END,
                       dst_sheet.name, e.dst_row_min, e.dst_row_max,
                       e.dst_col_min, e.dst_col_max, e.via
                FROM edges AS e
                JOIN sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
                LEFT JOIN fblocks AS src_block
                       ON e.src_kind = 'fblock' AND src_block.id = e.src_id
                LEFT JOIN sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
                ORDER BY src_sheet.name, e.src_kind, 3, COALESCE(dst_sheet.name, ''),
                         e.dst_row_min, e.dst_col_min, e.dst_row_max, e.dst_col_max, e.via
            """,
            "diagnostics": """
                SELECT d.severity, d.code, s.name, d.row, d.col, d.ref,
                       d.message, d.related
                FROM diagnostics AS d JOIN sheets AS s ON s.id = d.sheet_id
                ORDER BY s.name, COALESCE(d.row, -1), COALESCE(d.col, -1),
                         d.code, d.ref, d.message
            """,
            "staleness": """
                SELECT s.name, st.row_min, st.row_max, st.col_min, st.col_max, st.since
                FROM staleness AS st JOIN sheets AS s ON s.id = st.sheet_id
                ORDER BY s.name, st.row_min, st.col_min, st.row_max, st.col_max, st.since
            """,
            "cells": """
                SELECT s.name, c.row, c.col, c.ref, c.value, c.value_type,
                       c.formula, c.style_idx, c.formula_kind,
                       c.shared_index, c.array_ref, c.data_table
                FROM cells AS c JOIN sheets AS s ON s.id = c.sheet_id
                ORDER BY s.name, c.row, c.col
            """,
            "package_parts": """
                SELECT part_name, part_hash, kind
                FROM package_parts ORDER BY part_name
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
                       r.col_min, r.col_max
                FROM edge_rtree AS r
                JOIN edges AS e ON e.id = r.edge_id
                JOIN sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
                LEFT JOIN fblocks AS src_block
                       ON e.src_kind = 'fblock' AND src_block.id = e.src_id
                LEFT JOIN sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
                JOIN sheets AS spatial_min_sheet
                     ON spatial_min_sheet.id = CAST(r.sheet_min AS INTEGER)
                JOIN sheets AS spatial_max_sheet
                     ON spatial_max_sheet.id = CAST(r.sheet_max AS INTEGER)
                ORDER BY src_sheet.name, e.src_kind, 3, COALESCE(dst_sheet.name, ''),
                         e.dst_row_min, e.dst_col_min, e.dst_row_max, e.dst_col_max,
                         e.via, spatial_min_sheet.name, spatial_max_sheet.name,
                         r.row_min, r.col_min, r.row_max, r.col_max
            """
        else:
            queries["edge_intervals"] = """
                SELECT e.src_kind, src_sheet.name,
                       CASE WHEN e.src_kind = 'fblock' THEN src_block.n ELSE e.src_id END,
                       dst_sheet.name, e.dst_row_min, e.dst_row_max,
                       e.dst_col_min, e.dst_col_max, e.via,
                       spatial_sheet.name, i.row_min, i.row_max, i.col_min, i.col_max
                FROM edge_intervals AS i
                JOIN edges AS e ON e.id = i.edge_id
                JOIN sheets AS src_sheet ON src_sheet.id = e.src_sheet_id
                LEFT JOIN fblocks AS src_block
                       ON e.src_kind = 'fblock' AND src_block.id = e.src_id
                LEFT JOIN sheets AS dst_sheet ON dst_sheet.id = e.dst_sheet_id
                JOIN sheets AS spatial_sheet ON spatial_sheet.id = i.sheet_id
                ORDER BY src_sheet.name, e.src_kind, 3, COALESCE(dst_sheet.name, ''),
                         e.dst_row_min, e.dst_col_min, e.dst_row_max, e.dst_col_max,
                         e.via, spatial_sheet.name,
                         i.row_min, i.col_min, i.row_max, i.col_max
            """

        return {
            name: tuple(tuple(row) for row in self._connection.execute(sql).fetchall())
            for name, sql in queries.items()
        }

    def _initialize_schema(self, *, prefer_rtree: bool) -> None:
        self._connection.execute("PRAGMA foreign_keys = OFF")
        try:
            for attempt in range(len(_INITIALIZATION_RETRY_DELAYS) + 1):
                self.schema_created = False
                self.schema_rebuilt = False
                try:
                    self._connection.execute("BEGIN IMMEDIATE")
                    self._initialize_schema_locked(prefer_rtree=prefer_rtree)
                    self._connection.commit()
                    return
                except sqlite3.OperationalError as error:
                    if self._connection.in_transaction:
                        self._connection.rollback()
                    if not _is_busy_error(error) or attempt == len(_INITIALIZATION_RETRY_DELAYS):
                        raise
                    time.sleep(_INITIALIZATION_RETRY_DELAYS[attempt])
                except BaseException:
                    if self._connection.in_transaction:
                        self._connection.rollback()
                    raise
        finally:
            self._connection.execute("PRAGMA foreign_keys = ON")

    def _initialize_schema_locked(self, *, prefer_rtree: bool) -> None:
        if not self._table_exists("meta"):
            existing_tables = self._user_tables()
            if existing_tables:
                self._drop_all_objects()
                self.schema_rebuilt = True
            self._create_schema(initial_generation=0, prefer_rtree=prefer_rtree)
            self.schema_created = True
            return

        schema_version = self.get_meta("schema_version")
        if schema_version == SCHEMA_VERSION:
            EdgeStore.ensure_schema(self._connection, prefer_rtree=prefer_rtree)
            return

        previous_generation = self.generation
        self._drop_all_objects()
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

    def _drop_all_objects(self) -> None:
        if self._table_exists(EdgeStore.RTREE_TABLE):
            self._connection.execute("DROP TABLE edge_rtree")
        for object_type in ("view", "trigger"):
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
            "validations",
            "cells",
        ):
            if table == "columns":
                self._connection.execute(
                    "DELETE FROM columns WHERE region_id IN "
                    "(SELECT id FROM regions WHERE sheet_id = ?)",
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


def _normalize_part_name(part_name: str) -> str:
    return part_name.replace("\\", "/").lstrip("/")


def _part_kind(part_name: str) -> str:
    if part_name == "xl/workbook.xml":
        return "workbook"
    if part_name == "xl/_rels/workbook.xml.rels":
        return "workbook_rels"
    if part_name == "xl/sharedStrings.xml":
        return "shared_strings"
    if part_name == "xl/styles.xml":
        return "styles"
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
