"""SQLite storage and freshness lifecycle for workbook indexes."""

from excel_lsp.core.index.edges import EdgeBackend, EdgeStore
from excel_lsp.core.index.lifecycle import ensure_fresh, index_workbook, resolve_index_path
from excel_lsp.core.index.schema import SCHEMA_VERSION
from excel_lsp.core.index.store import IndexStore

__all__ = [
    "SCHEMA_VERSION",
    "EdgeBackend",
    "EdgeStore",
    "IndexStore",
    "ensure_fresh",
    "index_workbook",
    "resolve_index_path",
]
