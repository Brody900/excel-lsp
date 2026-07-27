"""SQLite schema constants for one workbook index."""

from __future__ import annotations

SCHEMA_VERSION = "3"

BASE_SCHEMA_SQL = """
CREATE TABLE meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE sheets (
    id INTEGER PRIMARY KEY,
    name TEXT UNIQUE NOT NULL,
    xml_part TEXT NOT NULL,
    part_hash TEXT NOT NULL,
    kind TEXT NOT NULL,
    visibility TEXT NOT NULL,
    max_row INTEGER NOT NULL,
    max_col INTEGER NOT NULL
);

CREATE TABLE regions (
    id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL,
    n INTEGER NOT NULL,
    row_min INTEGER NOT NULL,
    row_max INTEGER NOT NULL,
    col_min INTEGER NOT NULL,
    col_max INTEGER NOT NULL,
    header_rows INTEGER NOT NULL,
    kind TEXT NOT NULL,
    list_object_name TEXT,
    confidence REAL NOT NULL,
    FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE,
    UNIQUE (sheet_id, n)
);

-- Internal ListObject catalog. The public region schema deliberately stays
-- frozen, while formula analysis retains the header/totals metadata needed to
-- resolve structured references exactly after an incremental refresh.
CREATE TABLE list_objects (
    id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL,
    name TEXT NOT NULL,
    lookup_name TEXT NOT NULL UNIQUE,
    display_name TEXT NOT NULL,
    row_min INTEGER NOT NULL,
    row_max INTEGER NOT NULL,
    col_min INTEGER NOT NULL,
    col_max INTEGER NOT NULL,
    header_rows INTEGER NOT NULL,
    totals_rows INTEGER NOT NULL,
    FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE
);

CREATE TABLE list_object_columns (
    id INTEGER PRIMARY KEY,
    list_object_id INTEGER NOT NULL,
    idx INTEGER NOT NULL,
    name TEXT NOT NULL,
    lookup_name TEXT NOT NULL,
    FOREIGN KEY (list_object_id) REFERENCES list_objects(id) ON DELETE CASCADE,
    UNIQUE (list_object_id, idx),
    UNIQUE (list_object_id, lookup_name)
);

CREATE TABLE columns (
    id INTEGER PRIMARY KEY,
    region_id INTEGER NOT NULL,
    idx INTEGER NOT NULL,
    header TEXT NOT NULL,
    norm_header TEXT NOT NULL,
    dtype TEXT NOT NULL,
    nonnull INTEGER NOT NULL,
    distinct_est INTEGER NOT NULL,
    formula_block_id INTEGER,
    FOREIGN KEY (region_id) REFERENCES regions(id) ON DELETE CASCADE,
    UNIQUE (region_id, idx)
);

CREATE TABLE fblocks (
    id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL,
    n INTEGER NOT NULL,
    r1c1 TEXT NOT NULL,
    row_min INTEGER NOT NULL,
    row_max INTEGER NOT NULL,
    col_min INTEGER NOT NULL,
    col_max INTEGER NOT NULL,
    volatile INTEGER NOT NULL,
    opaque INTEGER NOT NULL,
    FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE,
    UNIQUE (sheet_id, n)
);

CREATE TABLE defined_names (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    scope_sheet_id INTEGER,
    refers_to TEXT NOT NULL,
    kind TEXT NOT NULL,
    is_builtin INTEGER NOT NULL,
    FOREIGN KEY (scope_sheet_id) REFERENCES sheets(id) ON DELETE CASCADE
);

CREATE TABLE name_areas (
    id INTEGER PRIMARY KEY,
    name_id INTEGER NOT NULL,
    sheet_id INTEGER NOT NULL,
    row_min INTEGER NOT NULL,
    row_max INTEGER NOT NULL,
    col_min INTEGER NOT NULL,
    col_max INTEGER NOT NULL,
    FOREIGN KEY (name_id) REFERENCES defined_names(id) ON DELETE CASCADE,
    FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE
);

CREATE TABLE validations (
    id INTEGER PRIMARY KEY,
    sheet_id INTEGER NOT NULL,
    row_min INTEGER NOT NULL,
    row_max INTEGER NOT NULL,
    col_min INTEGER NOT NULL,
    col_max INTEGER NOT NULL,
    vtype TEXT,
    operator TEXT,
    formula1 TEXT,
    formula2 TEXT,
    allow_blank INTEGER NOT NULL,
    FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE
);

CREATE TABLE edges (
    id INTEGER PRIMARY KEY,
    src_kind TEXT NOT NULL,
    src_id INTEGER NOT NULL,
    src_sheet_id INTEGER NOT NULL,
    dst_sheet_id INTEGER,
    dst_row_min INTEGER,
    dst_row_max INTEGER,
    dst_col_min INTEGER,
    dst_col_max INTEGER,
    via TEXT NOT NULL,
    FOREIGN KEY (src_sheet_id) REFERENCES sheets(id) ON DELETE CASCADE,
    FOREIGN KEY (dst_sheet_id) REFERENCES sheets(id) ON DELETE CASCADE
);

CREATE TABLE diagnostics (
    id INTEGER PRIMARY KEY,
    severity TEXT NOT NULL,
    code TEXT NOT NULL,
    sheet_id INTEGER NOT NULL,
    row INTEGER,
    col INTEGER,
    ref TEXT NOT NULL,
    message TEXT NOT NULL,
    related TEXT NOT NULL,
    FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE
);

CREATE TABLE staleness (
    sheet_id INTEGER NOT NULL,
    row_min INTEGER NOT NULL,
    row_max INTEGER NOT NULL,
    col_min INTEGER NOT NULL,
    col_max INTEGER NOT NULL,
    since TEXT NOT NULL,
    FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE
);

-- Internal parser stream. `value` has no affinity so normalized JSON scalars
-- retain SQLite's INTEGER/REAL/TEXT/NULL storage classes.
CREATE TABLE cells (
    sheet_id INTEGER NOT NULL,
    row INTEGER NOT NULL,
    col INTEGER NOT NULL,
    ref TEXT NOT NULL,
    value,
    value_type TEXT NOT NULL,
    formula TEXT,
    style_idx INTEGER NOT NULL,
    formula_kind TEXT,
    shared_index INTEGER,
    array_ref TEXT,
    data_table TEXT,
    PRIMARY KEY (sheet_id, row, col),
    FOREIGN KEY (sheet_id) REFERENCES sheets(id) ON DELETE CASCADE
) WITHOUT ROWID;

-- Selected OOXML hashes used to choose a full, per-sheet, or no-op refresh.
CREATE TABLE package_parts (
    part_name TEXT PRIMARY KEY,
    part_hash TEXT NOT NULL,
    kind TEXT NOT NULL
);

CREATE INDEX regions_bounds
    ON regions(sheet_id, row_min, row_max, col_min, col_max);
CREATE INDEX list_objects_bounds
    ON list_objects(sheet_id, row_min, row_max, col_min, col_max);
CREATE INDEX fblocks_bounds
    ON fblocks(sheet_id, row_min, row_max, col_min, col_max);
CREATE INDEX name_areas_bounds
    ON name_areas(sheet_id, row_min, row_max, col_min, col_max);
CREATE INDEX validations_bounds
    ON validations(sheet_id, row_min, row_max, col_min, col_max);
CREATE INDEX diagnostics_location ON diagnostics(sheet_id, row, col);
CREATE INDEX staleness_bounds
    ON staleness(sheet_id, row_min, row_max, col_min, col_max);
CREATE INDEX cells_formula ON cells(sheet_id, formula) WHERE formula IS NOT NULL;
"""

CONTENT_TABLES_DELETE_ORDER = (
    "diagnostics",
    "staleness",
    "edges",
    "columns",
    "regions",
    "fblocks",
    "list_object_columns",
    "list_objects",
    "name_areas",
    "defined_names",
    "validations",
    "cells",
    "sheets",
)


__all__ = ["BASE_SCHEMA_SQL", "CONTENT_TABLES_DELETE_ORDER", "SCHEMA_VERSION"]
