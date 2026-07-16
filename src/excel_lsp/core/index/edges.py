"""Spatial range storage with an SQLite R*Tree and interval-table fallback."""

from __future__ import annotations

import sqlite3
from typing import Literal

from excel_lsp.core.models import Rect

EdgeBackend = Literal["rtree", "interval"]


class EdgeStore:
    """Store/query edge destination rectangles without exposing the backend.

    R*Tree is preferred. SQLite builds without the module use a plain table and
    the same inclusive interval-overlap predicates. Keeping the fallback here
    prevents graph callers from acquiring backend-specific SQL.
    """

    RTREE_TABLE = "edge_rtree"
    INTERVAL_TABLE = "edge_intervals"

    def __init__(self, connection: sqlite3.Connection) -> None:
        self._connection = connection
        if self._table_exists(self.RTREE_TABLE):
            self.backend: EdgeBackend = "rtree"
        elif self._table_exists(self.INTERVAL_TABLE):
            self.backend = "interval"
        else:
            raise RuntimeError("edge range schema has not been initialized")

    @classmethod
    def ensure_schema(
        cls, connection: sqlite3.Connection, *, prefer_rtree: bool = True
    ) -> EdgeBackend:
        """Create one spatial backend and return the selected implementation."""
        tables = {
            str(row[0])
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
        }
        if cls.RTREE_TABLE in tables:
            return "rtree"
        if cls.INTERVAL_TABLE in tables:
            return "interval"

        if prefer_rtree:
            try:
                connection.execute(
                    "CREATE VIRTUAL TABLE edge_rtree USING rtree("
                    "edge_id, sheet_min, sheet_max, row_min, row_max, col_min, col_max)"
                )
                return "rtree"
            except sqlite3.OperationalError as exc:
                # Only module availability is recoverable. Syntax, disk, and
                # corruption failures must not be silently hidden.
                message = str(exc).casefold()
                if "no such module" not in message or "rtree" not in message:
                    raise

        connection.execute(
            """
            CREATE TABLE edge_intervals (
                edge_id INTEGER PRIMARY KEY,
                sheet_id INTEGER NOT NULL,
                row_min INTEGER NOT NULL,
                row_max INTEGER NOT NULL,
                col_min INTEGER NOT NULL,
                col_max INTEGER NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE INDEX edge_intervals_overlap
            ON edge_intervals(sheet_id, row_min, row_max, col_min, col_max)
            """
        )
        return "interval"

    def clear(self) -> None:
        """Remove all indexed destination rectangles."""
        self._connection.execute(f"DELETE FROM {self.table_name}")

    def insert(self, edge_id: int, sheet_id: int, rect: Rect) -> None:
        """Insert or replace one inclusive destination rectangle."""
        if self.backend == "rtree":
            self._connection.execute(
                "INSERT OR REPLACE INTO edge_rtree("
                "edge_id, sheet_min, sheet_max, row_min, row_max, col_min, col_max"
                ") VALUES (?, ?, ?, ?, ?, ?, ?)",
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
            return
        self._connection.execute(
            "INSERT OR REPLACE INTO edge_intervals("
            "edge_id, sheet_id, row_min, row_max, col_min, col_max"
            ") VALUES (?, ?, ?, ?, ?, ?)",
            (
                edge_id,
                sheet_id,
                rect.row_min,
                rect.row_max,
                rect.col_min,
                rect.col_max,
            ),
        )

    def delete(self, edge_id: int) -> None:
        """Delete one destination rectangle if present."""
        self._connection.execute(f"DELETE FROM {self.table_name} WHERE edge_id = ?", (edge_id,))

    def query_point(self, sheet_id: int, row: int, col: int) -> tuple[int, ...]:
        """Return edge ids whose destinations contain a worksheet point."""
        if self.backend == "rtree":
            rows = self._connection.execute(
                """
                SELECT edge_id FROM edge_rtree
                WHERE sheet_min <= ? AND sheet_max >= ?
                  AND row_min <= ? AND row_max >= ?
                  AND col_min <= ? AND col_max >= ?
                ORDER BY edge_id
                """,
                (sheet_id, sheet_id, row, row, col, col),
            )
        else:
            rows = self._connection.execute(
                """
                SELECT edge_id FROM edge_intervals
                WHERE sheet_id = ?
                  AND row_min <= ? AND row_max >= ?
                  AND col_min <= ? AND col_max >= ?
                ORDER BY edge_id
                """,
                (sheet_id, row, row, col, col),
            )
        return tuple(int(item[0]) for item in rows.fetchall())

    def query_range(self, sheet_id: int, rect: Rect) -> tuple[int, ...]:
        """Return edge ids whose destinations overlap an inclusive rectangle."""
        if self.backend == "rtree":
            rows = self._connection.execute(
                """
                SELECT edge_id FROM edge_rtree
                WHERE sheet_min <= ? AND sheet_max >= ?
                  AND row_min <= ? AND row_max >= ?
                  AND col_min <= ? AND col_max >= ?
                ORDER BY edge_id
                """,
                (
                    sheet_id,
                    sheet_id,
                    rect.row_max,
                    rect.row_min,
                    rect.col_max,
                    rect.col_min,
                ),
            )
        else:
            rows = self._connection.execute(
                """
                SELECT edge_id FROM edge_intervals
                WHERE sheet_id = ?
                  AND row_min <= ? AND row_max >= ?
                  AND col_min <= ? AND col_max >= ?
                ORDER BY edge_id
                """,
                (
                    sheet_id,
                    rect.row_max,
                    rect.row_min,
                    rect.col_max,
                    rect.col_min,
                ),
            )
        return tuple(int(item[0]) for item in rows.fetchall())

    @property
    def table_name(self) -> str:
        """Return the physical table used by this connection."""
        return self.RTREE_TABLE if self.backend == "rtree" else self.INTERVAL_TABLE

    def _table_exists(self, table_name: str) -> bool:
        row = self._connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
            (table_name,),
        ).fetchone()
        return row is not None


__all__ = ["EdgeBackend", "EdgeStore"]
