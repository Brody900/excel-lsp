"""SQLite schema constants for one workbook index."""

from __future__ import annotations

SCHEMA_VERSION = "5"

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
    dependent_rank INTEGER,
    precedent_rank INTEGER,
    FOREIGN KEY (src_sheet_id) REFERENCES sheets(id) ON DELETE CASCADE,
    FOREIGN KEY (dst_sheet_id) REFERENCES sheets(id) ON DELETE CASCADE
);

-- Ranked spatial mirrors are a derived graph index.  Mutations to any table
-- contributing public graph labels or source geometry make them unavailable
-- until the store atomically rebuilds and validates both directions.
CREATE TABLE graph_spatial_state (
    singleton INTEGER PRIMARY KEY CHECK (singleton = 1),
    dirty INTEGER NOT NULL CHECK (dirty IN (0, 1)),
    dependent_rank_max INTEGER NOT NULL,
    precedent_rank_max INTEGER NOT NULL,
    revision INTEGER NOT NULL,
    mutation_epoch INTEGER NOT NULL,
    clean_epoch INTEGER NOT NULL
);
INSERT INTO graph_spatial_state VALUES (1, 1, 0, 0, 0, 0, 0);

-- Canonical public-hop identity for every dense rank in both directions.
-- Graph-affecting mutations delete impacted identities, so restoring only
-- the seven-field graph seal cannot make a split rank trustworthy again.
CREATE TABLE graph_rank_keys (
    direction TEXT NOT NULL CHECK (direction IN ('dependents', 'precedents')),
    rank INTEGER NOT NULL CHECK (rank > 0),
    key_text TEXT NOT NULL,
    PRIMARY KEY (direction, rank)
) WITHOUT ROWID;

CREATE TRIGGER edges_graph_spatial_dirty_insert AFTER INSERT ON edges BEGIN
    DELETE FROM graph_rank_keys
    WHERE (direction = 'dependents' AND rank = NEW.dependent_rank)
       OR (direction = 'precedents' AND rank = NEW.precedent_rank);
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER edges_graph_spatial_dirty_update AFTER UPDATE ON edges BEGIN
    DELETE FROM graph_rank_keys
    WHERE (direction = 'dependents'
           AND rank IN (OLD.dependent_rank, NEW.dependent_rank))
       OR (direction = 'precedents'
           AND rank IN (OLD.precedent_rank, NEW.precedent_rank));
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER edges_graph_spatial_dirty_delete AFTER DELETE ON edges BEGIN
    DELETE FROM graph_rank_keys
    WHERE (direction = 'dependents' AND rank = OLD.dependent_rank)
       OR (direction = 'precedents' AND rank = OLD.precedent_rank);
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER fblocks_graph_spatial_dirty_insert AFTER INSERT ON fblocks BEGIN
    DELETE FROM graph_rank_keys;
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER fblocks_graph_spatial_dirty_update AFTER UPDATE ON fblocks BEGIN
    DELETE FROM graph_rank_keys;
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER fblocks_graph_spatial_dirty_delete AFTER DELETE ON fblocks BEGIN
    DELETE FROM graph_rank_keys;
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER sheets_graph_spatial_dirty_insert AFTER INSERT ON sheets BEGIN
    DELETE FROM graph_rank_keys;
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER sheets_graph_spatial_dirty_update AFTER UPDATE ON sheets BEGIN
    DELETE FROM graph_rank_keys;
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER sheets_graph_spatial_dirty_delete AFTER DELETE ON sheets BEGIN
    DELETE FROM graph_rank_keys;
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;

CREATE TRIGGER graph_rank_keys_graph_spatial_dirty_insert
AFTER INSERT ON graph_rank_keys BEGIN
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER graph_rank_keys_graph_spatial_dirty_update
AFTER UPDATE ON graph_rank_keys BEGIN
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;
CREATE TRIGGER graph_rank_keys_graph_spatial_dirty_delete
AFTER DELETE ON graph_rank_keys BEGIN
    UPDATE graph_spatial_state
    SET dirty = 1, mutation_epoch = mutation_epoch + 1
    WHERE singleton = 1;
END;

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
CREATE INDEX edges_source ON edges(src_sheet_id, src_kind, src_id, via);
CREATE INDEX edges_dependent_rank ON edges(dependent_rank);
CREATE INDEX edges_precedent_rank ON edges(precedent_rank);
CREATE INDEX edges_precedent_semantic
    ON edges(
        dst_sheet_id, src_sheet_id,
        dst_row_min, dst_col_min, dst_row_max, dst_col_max,
        via, src_kind, src_id
    );
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
