# Index internals

Excel LSP stores one disposable SQLite semantic index per workbook. P1's parser,
schema, persistence, spatial-backend abstraction, canonical export, and
freshness lifecycle are verified. P2's region, column, symbol, and compact-map
contracts are also verified. Tables for later graph, diagnostics, and editor
phases are reserved schema, not evidence that those features already exist.

## Sidecar placement

By default a workbook at `model.xlsx` uses `model.xlsx.xlsp.db` beside the
source file. Setting `EXCEL_LSP_INDEX_DIR` places the sidecar in that directory
as `<stem>.<path-sha256-prefix>.xlsp.db`, avoiding collisions between equal
filenames from different directories.

The workbook remains authoritative. A sidecar may be deleted while Excel LSP is
not using it and will be rebuilt. Sidecars, WAL/SHM companions, and generated
fixture workbooks are ignored by Git.

## Connection and transaction model

Every `IndexStore` connection configures:

- a 5-second busy timeout;
- WAL journal mode;
- `synchronous=NORMAL`;
- foreign-key enforcement;
- explicit `BEGIN IMMEDIATE` mutations with rollback on failure.

Cold schema creation and migration serialize through the same immediate
transaction with bounded busy retries. The schema version is stored in both
`meta` and SQLite's `user_version`. A mismatch drops and recreates the derived
schema rather than attempting an in-place data migration.

`meta.generation` is a monotonically increasing semantic revision. It changes
on index mutations, not when only file-stat bookkeeping changes. P7 pagination
cursors will embed it so a write or reindex cannot produce a mixed-generation
response.

## Schema ownership

| Table | Owner and purpose | Current status |
|---|---|---|
| `meta` | Schema, source path/hash/stat, generation, date system, VBA and external-link metadata | Verified P1 |
| `package_parts` | Selected normalized OOXML part hashes and part kinds | Verified P1 |
| `sheets` | Workbook-order catalog, part hashes, kind, visibility and actual bounds | Verified P1 |
| `cells` | Sparse normalized cell stream with formulas, styles, shared/array/Data Table metadata | Verified P1 |
| `defined_names`, `name_areas` | Global/sheet names and resolvable area rectangles | Verified P1 |
| `validations` | Sparse validation rectangles and constraints | Verified P1 |
| `regions`, `columns` | Region bounds, headers, type summaries and confidence | Verified P2 |
| `fblocks` | R1C1 formula-block rectangles and flags | P3 planned |
| `edges` plus spatial table | Formula destinations and overlap lookup | Storage verified P1; graph P4 planned |
| `diagnostics` | Bounded workbook and formula findings | P5 planned |
| `staleness` | Rectangles affected by surgical writes | P6 planned |

`cells` is a `WITHOUT ROWID` table keyed by `(sheet_id, row, col)`. Values cross
the shared JSON-scalar normalization boundary before persistence. Excel numeric
values outside SQLite's signed 64-bit integer range are stored as finite REALs
rather than causing a raw driver overflow.

Foreign keys cascade derived rows when a sheet or region is removed. Per-sheet
refresh deliberately deletes only rows owned by that source sheet; incoming
edges from other sheets are retained because their formulas did not change.

## Spatial range backend

The `EdgeStore` exposes one inclusive rectangle interface for both physical
backends:

- preferred `edge_rtree(edge_id, sheet_min, sheet_max, row_min, row_max,
  col_min, col_max)`;
- fallback `edge_intervals(edge_id, sheet_id, row_min, row_max, col_min,
  col_max)` with an overlap index.

Point containment and range overlap return ordered edge IDs with the same
semantics. R*Tree initialization falls back only when SQLite explicitly reports
that the module is unavailable; corruption or SQL errors are not hidden as a
portability fallback.

Canonical exports project spatial rows through natural sheet names and natural
edge identities. Physical sheet, edge, region, and other surrogate IDs are
excluded so independently built full and incremental databases can be compared
meaningfully.

## Refresh lifecycle

Each core entry point can call the same freshness path:

1. Resolve and validate the workbook suffix and source stat.
2. If stored path, `mtime_ns`, and size match, return without parsing or changing
   generation.
3. Otherwise stream workbook metadata and compare the whole-file and selected
   part hashes.
4. If the whole-file hash is unchanged, repair stat metadata and preserve
   generation.
5. Workbook structure or content-type changes rebuild the catalog and all
   sheets. Shared-string or style changes refresh every sheet. Worksheet,
   relationship, table-part, merge, or validation changes refresh only the
   owning sheet.
6. Re-stat before commit. If Excel saved concurrently, roll back and retry once;
   repeated change becomes a structured corrupt/torn-save error.
7. Atomically replace hashes and metadata and bump generation exactly once for
   a semantic mutation.

This implements the P1 invariants: untouched refresh is a no-op, independent
full and incremental builds have equal natural-key exports, and indexing never
modifies the source bytes. The executable and review evidence is in
[`docs/evidence/p1-foundation.md`](evidence/p1-foundation.md).

## P2 regions and headers contract

P2 must populate `regions` and `columns` without materializing a dense sheet:

- stream non-empty column intervals per row and merge them vertically with a
  configurable one-row/one-column default gap tolerance;
- emit parsed ListObjects first and prevent heuristic regions from overlapping
  their exact declared rectangles;
- inspect at most three candidate header rows using string/body type shifts,
  uniqueness, merged cells, and lazily resolved bold/fill features;
- sample at most 200 body cells per column for
  `int`, `float`, `date`, `str`, `bool`, `mixed`, or `empty` classification;
- count non-null cells and cap distinct-value tracking at 1,000 hashes;
- join merged multi-row header text deterministically;
- reduce sampling and emit `W_LARGE_SHEET` above two million non-empty cells;
- order regions by top row and left column so ordinals are deterministic.

The implementation keeps those rules sparse under adversarial metadata. A
spatially partitioned tree selects ListObject barriers while preserving exact
rect-key result order; only a consecutive full-height barrier prefix is
partitioned in one run pass. Each anchored merged range stays one raw span
rather than expanding once per covered row. Ordinary runs are coalesced once.
On merge-bearing sheets, a later-primitive sweep first forms table-aware sparse
components and closes intersecting component bounds; only then does the
fixed-order table BSP clip a component whose own bound intersects that
ListObject. Each directional child is recomputed independently, so siblings
cannot reconnect across a barrier.

Two primitives within the configured blank-row and blank-column gaps connect
only when at least one minimal cell-to-cell witness rectangle is table-free.
Horizontal and vertical corridors use exact projected-interval coverage;
diagonal corridors use their nearest boundary rectangle. The sweep keeps an
all-root rectangle index for bounding-box closure and a row-active root index
for proximity. Each root owns power-of-two immutable member-index blocks that
meld binomially: live membership remains linear, each member is rebuilt only
logarithmically, and queries do not enumerate unrelated or already-internal
members. The ordinary no-merge path retains its simpler row-run fast path.
Merged headers use at-most-three-row interval views. Committed work-count and
brute differential tests cover each optimization.

`gap_tol` defaults to 1 and accepts the bounded range 0–8. The upper bound keeps
the active-row component work linear and prevents a configuration value from
turning sparse detection into an unbounded all-prior-row comparison. Invalid
values fail before an index file is created.

The phase invariants are exact: each non-empty cell belongs to at most one
region, every ListObject range is reproduced exactly, and identical content
produces identical region ordinals. Those statements remain acceptance criteria
until F02/F12/F20 tests and P2 evidence prove them.

## Frozen symbol identifiers

```text
sheet:{sheetName}
region:{sheetName}:{n}
col:{sheetName}:{n}:{normalizedHeader}[#k]
name:{definedName}
name:{sheetName}!{definedName}
fblock:{sheetName}:{n}
cell:{sheetName}!{A1}
```

Region and formula-block ordinals are zero-based. Column headers use Unicode
NFKC normalization, case folding, separator collapsing, deterministic fallbacks,
and a `#k` suffix for duplicates. Region IDs are stable when region content and
ordering are unchanged, but may shift when regions are added or moved; callers
should cheaply fetch a new workbook map after refresh.

## Compact workbook map contract

The P2 map projection reads only semantic summaries from the index. It includes
the workbook name, sheet count, index timestamp, stale/VBA flags, sheet
visibility and actual dimensions, bounded regions and column summaries, defined
name references, external-link labels, diagnostic counts, and navigation hints.
It includes no raw cell values.

Normal degradation limits are eight regions per sheet, 20 names, and 10
external links, with explicit remainder counts. Additional deterministic
degradation may reduce columns, regions, names, links, and finally detailed
sheets until serialized output is within 8,000 characters. Hidden and
very-hidden sheets are prioritized when sheet detail itself must be reduced.
The F03 reference map has the separate 1,500-token `o200k_base` budget.

The detector, persistence integration, bounded projection, symbol helpers,
deterministic fixtures, golden output, and budget measurements are verified
through P2's R-mech, R-test, and user-authorized early R-repo gates. Current
artifacts are indexed in [the P2 evidence report](evidence/p2-regions-map.md);
the release proof paths are listed in the
[README claims-to-artifacts plan](evidence/readme-claims-to-artifacts.md).

## Formula blocks

Planned for P3. This section will document A1-to-R1C1 normalization, contiguous
block construction, sheet-bounds-clamped extrusion, modern-formula handling,
and inconsistency detection after their executable evidence exists.

## Range edges

Planned for P4. This section will document graph population, rectangle edge
semantics, precedent/dependent traversal, path queries, and bounded circular
detection after their executable evidence exists. P1 verified only the physical
R*Tree/interval abstraction described above.

## Later-phase population

P3 populates `fblocks`; P4 populates graph edges and spatial rectangles; P5
populates diagnostics; P6 populates staleness and applies direct post-write index
patches. P7 exposes these through bounded MCP and CLI calls. Until each phase
gate closes, the presence of its schema table is only forward-compatible
storage design.
