# Index internals

Excel LSP stores one disposable SQLite semantic index per workbook. P1's parser,
schema, persistence, spatial-backend abstraction, canonical export, and
freshness lifecycle are verified. P2's region, column, symbol, and compact-map
contracts are also verified. Verified P3 populates formula blocks, reference
edges, ListObject context, and its bounded subset of formula diagnostics.
Verified P4 adds ranked spatial graph navigation and bounded circular
analysis. Verified P5 completes the typed diagnostic catalog, persisted
cached-error/link/volatility findings, and filtered query. Editor tables remain
reserved schema, not evidence that those later features exist.

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
| `list_objects`, `list_object_columns` | Durable ListObject aliases, bounds, header/totals counts, and ordered column names used by structured-reference analysis | Verified P3 |
| `fblocks` | Exact R1C1 formula-block rectangles, flags, and formula count | Verified P3 |
| `edges`, ranked source/destination mirrors, `graph_spatial_state`, `graph_rank_keys` | P3 formula destinations plus P4 exact-order spatial traversal, bounded rank identity, and mirror integrity | Verified P4 |
| `diagnostics` | P2 large-sheet, P3 parse/name/dynamic/inconsistency, P4 circular, and P5 cached-error/link/volatile findings | Verified P5 |
| `staleness` | Rectangles affected by surgical writes | P6 planned |

`cells` is a `WITHOUT ROWID` table keyed by `(sheet_id, row, col)`. Values cross
the shared JSON-scalar normalization boundary before persistence. Excel numeric
values outside SQLite's signed 64-bit integer range are stored as finite REALs
rather than causing a raw driver overflow.

Foreign keys cascade derived rows when a sheet or region is removed. Per-sheet
refresh deliberately deletes only rows owned by that source sheet; incoming
edges from other sheets are retained because their formulas did not change.

## Spatial range backend

The `EdgeStore` exposes inclusive point/range overlap and ranked graph traversal
through equivalent physical backends:

- preferred destination `edge_rtree` and source `edge_source_rtree` virtual
  tables use `rtree_i32`; each indexes sheet, row, column, and one dense public
  hop rank;
- fallback `edge_intervals` and `edge_source_intervals` tables retain the same
  rectangles and ranks with ordinary overlap indexes.

Point containment and range overlap retain ordered edge-ID compatibility for
the P1/P3 storage API. Bounded P4 queries instead binary-probe the RTree rank
dimension to find successive intersecting public hops. This returns the exact
semantic prefix without sorting or materializing all matches. The interval
fallback builds a deterministic, revision-guarded rectangle BVH from the
ranked tables and uses minimum-rank branch-and-bound traversal. It therefore
retains exact parity and bounded irrelevant-edge work without relying on an
ordinary B-tree's first matching rank.

Every formula refresh recomputes dense dependent and precedent ranks over unique
public hops and atomically rebuilds both mirrors. Edge, block, sheet, or mirror
mutation marks `graph_spatial_state` dirty and increments `mutation_epoch`; a
successful rebuild seals the same value in `clean_epoch`. Graph calls require
the clean bit and equal epochs, so flipping one persisted bit cannot conceal a
missing mirror row. Current-schema opening validates the exact state-table
shape, exact rank-key catalog identity, complete canonical SQL for all 18 dirty
triggers, and rejects any extra
persistent trigger on protected graph tables or one that references graph trust
state. One shared projector both builds ranks and rederives every persisted
fblock/cell/range/opaque semantic key at open, proving exact case-aware dense
order rather than only rank density. Opening also compares the complete source
and destination mirrors in both directions, validates exact physical DDL, and
checks active RTree node/rowid storage. A process-local immutable trust tuple
seals all seven persisted state fields: singleton, dirty, both rank maxima,
revision, mutation epoch, and clean epoch. A second process-local live seal
records a graph-specific monotonic write epoch, other-connection `data_version`,
and `schema_version`; the complete canonical rank-key map is cached as immutable
selected-rank identity. `IndexStore` uses a private SQLite connection subclass
whose authorizer dispatcher increments that epoch for every allowed graph-table,
mirror, trigger, or relevant schema mutation while still honoring caller
authorizers. `IndexStore.connection` exposes a narrow SQL capability rather than
that native handle. It wraps every returned cursor, keeps `cursor.connection`
inside the same capability, routes caller authorizers through the tracker, and
rejects custom cursor factories and row-factory callbacks that would receive a
native cursor. Native SQLite base descriptors therefore cannot displace the
tracker through the supported public connection surface. Statement caching is
disabled so repeated SQL, cursor calls, and `executemany` cannot bypass the
dispatcher. Live query validation compares the
constant-size tokens in O(1) and consumes the already-validated maximum without
a second mutable read. A successful managed rebuild publishes its captured
state, catalog, and live seals verbatim after commit; it never rebases the epoch
over later graph writes. Rollback may rebase only after SQLite restores state.
For RTree, rollback first reconnects both virtual tables before that rebase;
SQLite can otherwise replay write-shaped authorizer callbacks while lazily
reloading a rolled-back TEMP schema even though `total_changes`, graph state,
and the rank catalog did not change.
One `graph_rank_keys` row stores the canonical serialized hop key for each dense
`(direction, rank)`. Its composite primary key makes selected-rank validation
O(1). Rebuild populates it after all active mirrors; edge, fblock, sheet,
catalog, RTree-shadow, and interval-mirror triggers invalidate the affected
rank or direction. Consequently, restoring the seven state fields, catalog
rows, and relational data still cannot make a split duplicate rank trustworthy:
same-handle restoration advances the private graph epoch, cross-handle
restoration changes `data_version`, and trigger replacement changes
`schema_version`. Non-graph metadata, diagnostics, and staleness writes leave the
graph epoch untouched. Open validation compares the complete catalog with
independently rederived canonical records, while bounded queries compare their
representative hop with the cached selected identity. Validation, seal capture,
and interval-cache warming occur inside one constructor-owned `BEGIN IMMEDIATE`
snapshot, preventing a writer commit from landing between those steps in WAL or
DELETE mode. Valid duplicate edges retain one catalog row and one public hop.
Each complete public graph operation also owns one deferred SQLite read
snapshot when the caller has not already opened a transaction. Trust
validation is the snapshot's first read, so later relational and spatial reads
cannot mix generations if another process commits concurrently. Graph-owned
snapshots are released on success and every exception through rollback-state
inspection and direct-SQL fallback. Snapshot `BEGIN` itself is inside this
cleanup boundary, because a native connection subclass may take effect and then
raise. If release remains unverifiable, the graph is poisoned and the supported
native SQLite descriptor is conclusively closed, including an emergency
base-descriptor path that bypasses failing subclass overrides. Native closure
state is read through the base `sqlite3.Connection.in_transaction` descriptor,
not a subclass-overridable property, so neither a false closed-style error nor
a false open-style error can misclassify the physical handle. Raw and
store-managed caller transactions remain entirely caller-owned.

Managed and constructor-owned `BEGIN IMMEDIATE` acquisition have the same
after-effect protection.
Failed transaction cleanup proves rollback or conclusively closes the supported
native descriptor, always finalizes process-local graph bookkeeping, preserves
the exact primary error, and retains earlier cause/context plus cleanup evidence
without causal cycles or repeated exception identities, including when the
primary is nested in an immutable cleanup group or two members share one
external causal chain. Group normalization rebuilds membership only when
needed, claims every recursively nested member first, assigns each external
causal identity one deterministic owner, and removes later cause/context
aliases, including suppressed contexts. The same closure-aware composition is
applied to successful snapshot multi-cleanup, direct close aggregation, and
post-commit hook/finalizer failures.
The normalizer is a returning sanitizer: immutable exception groups are derived
when duplicate or excluded members must be removed, and every caller installs
the returned root. It preserves message/subclass identity, notes, traceback,
cause, context, and the explicit suppression flag while assigning every member
or external causal object one owner. This applies equally when one group is the
only successful-query cleanup, close, post-commit hook, or constructor-release
failure.
`IndexStore.close()`
applies the same fallback discipline,
runs the no-I/O graph finalizer unconditionally, retries virtual close, bypasses
a failing native subclass override through `sqlite3.Connection.close`, and
marks the store closed only after physical closure is proven. Context exit keeps
the body error primary under the same causal-evidence rules. Constructor and
connection-configuration failures use the same native rollback/close proof even
before `edge_store` exists. Constructor cleanup owns the handle before the
public capability is allocated; early and final configuration-step failures,
virtual cleanup failures, and failures before or after rollback/close takes
effect all retain the exact initialization primary and unique causal evidence.
The native descriptor is closed unconditionally so an acquired writer lock
cannot leak.
Caller-owned raw SQLite transactions cannot be adopted for graph rebuilds
and are refused before mutation; use `IndexStore.transaction()` for managed
commit, rollback, and nesting.

Invalid state rebuilds the disposable sidecar monotonically. If corrupt virtual
storage prevents valid DDL teardown, the store builds and checkpoints a complete
same-directory temporary database, removes stale WAL/SHM files, and installs it
with atomic `os.replace`. Failed replacement preserves the original sidecar and
error and cleans temporary artifacts. Both the old and completed-replacement
handles use native rollback/closure proof, including post-commit virtual-close
failures. The `mkstemp` descriptor is cleanup-owned from acquisition, so a
close failure before or after native effect cannot leak the descriptor or its
`.rebuild` artifact. Graph calls otherwise return structured
`E_CORRUPT` rather than stale topology. Every authoritative graph read, write,
mirror join, catalog lookup, state check, and schema PRAGMA is explicitly bound
to `main`. Construction, schema setup, and managed mutation also reject any
TEMP object whose name or target can shadow protected graph storage; creating
one after facade construction changes either the tracked same-handle epoch or
the constant-size `temp.schema_version` seal, so tracked and plain-connection
facades both fail closed. Canonical export binds every projected table to
`main` and rejects protected TEMP shadows without requiring an otherwise-clean
intermediate graph. Direct writable-schema updates to `sqlite_master` and its
aliases are graph-authoritative too. The caller authorizer runs first, so a
denied graph write neither mutates SQLite nor advances the epoch. A caller
`SQLITE_IGNORE` verdict can let SQLite continue a write with altered semantics,
so graph-affecting ignored actions advance the private epoch before that
verdict is returned. R*Tree
initialization falls back only when SQLite explicitly reports
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

Verified P3 normalizes every formula cell independently with the
tokenizer-backed `to_r1c1`. Absolute row and column coordinates stay absolute;
relative coordinates become offsets from that cell. Names, structured
references, spill operands, quoted qualifiers, and non-reference tokens pass
through without being mistaken for ordinary A1 coordinates. Shared-formula
translation uses the same quote-, bracket-, and 3-D-aware reference grammar,
including modern `@` and `#` endpoint forms.

Block construction scans exact formula coordinates column-major, joins equal
contiguous R1C1 runs vertically, then merges only adjacent columns with exactly
matching row spans and signatures. Singleton and malformed-but-contained
formula patterns remain valid blocks. Consequently, every formula cell has one
owner and the union of block rectangles is exactly the formula-cell set.

References are classified once per block. Relative destination rectangles are
extruded across the block, absolute dimensions remain fixed, and every result
is clamped to Excel's row and column bounds. Structured operands whose meaning
depends on the source cell—including endpoints nested in colon, intersection,
or union expressions—are projected over an exact row-event tiling of
homogeneous ListObject contexts. The block-anchor formula is translated to each
tile anchor with coordinate spill anchors preserved before contextual
reanalysis. The block identity stays whole while edge and opacity coverage
remain equivalent to per-cell classification. Spill edges target only their
statically known definition anchor and are never extruded into a guessed
dynamic extent.

Reference-result inference is typed through lexical LET/LAMBDA bindings,
first-class named LAMBDAs, higher-order arguments and returned callables, and
conservative `CHOOSE`/`IF` callable alternatives. Computed range endpoints are
visible as `opaque:<callable>` plus `I_DYNAMIC_REF` unless a supported static
operator can be represented exactly. Static colon/intersection folding handles
transparent parentheses, whole axes, concrete range names, and compatible
structured endpoints; an unrepresentable mixed-coordinate intersection
remains explicitly opaque instead of silently claiming completeness. Constants
and formula/LAMBDA names are never treated as exact geometry merely because
their bodies expose one precedent. Their precedents remain indexed, while
result-reference metadata preserves `INDEX`/dynamic callable attribution
through LET, aliases, colon, and whitespace intersection. Expansion uses a
execution-local cache owned by retained `Thread` and asyncio `Task` objects.
Its identity key retains the exact `ReferenceContext`, anchor, spelling, spill
mode, and recursion stack so copied contexts and recycled numeric thread ids
cannot share mutable expansion state.

Excel-authored OOXML prefixes LET/LAMBDA declarations and their local uses with
`_xlpm.`. Lookup preserves that prefix as a distinct lexical namespace, while
public labels and duplicate-name checks use the prefix-free, case-insensitive
display spelling. Raw formula declarations follow the installed worksheet UI:
valid in-grid A1 spellings are rejected, while R1C1-like and beyond-grid
spellings are accepted. Stored `_xlpm.` declarations are authoritative even
when their suffix resembles a cell address. A lexical local participating in
`:` or whitespace intersection keeps every concrete precedent but also emits
conservative opacity; scalar locals never gain false reference identity, while
reference-valued locals retain outer dynamic attribution.

Formula inconsistency detection examines maximal vertical and horizontal runs
of at least five formulas. It reports minority signatures only when the
dominant signature covers at least 80% and the minority count is at most
`max(3, ceil(0.05*n))`, de-duplicating cells seen in both passes. F07 freezes the
one planted `C12` finding.

Formula analysis is part of the same index transaction as sheet/catalog
refresh. A worksheet change replaces only source-owned blocks, edges, and P3
diagnostics while retaining valid incoming edges; name, ListObject, or external
link context changes reclassify every affected source. A failure rolls back the
semantic rows and generation together. The executable candidate evidence is
in [the P3 report](evidence/p3-formulas-blocks.md).

## Range edges

P3 populates one or more destination rectangles per formula block for statically
resolved A1, 3-D, name, structured, spill-anchor, and supported composite
references. Dynamic/external/unresolved destinations retain a deliberate
`via` label and may have no destination rectangle. The R*Tree and interval
fallback persist equal natural-key semantics.

Verified P4 consumes these rows for direct precedents/dependents, bounded
breadth-first traces, shortest dependent paths, and two-stage circular
analysis. Candidate tests prove exact semantic-prefix ordering independently of
edge IDs, whole-column overlap, mirror corruption handling, and bounded work on
50,000 irrelevant edges through both spatial backends. A populated `edges`
table alone still proves only P3 extraction unless the verified graph schema,
mirrors, seals, and bounded query contracts are also present.

Transaction rollback bookkeeping is an internal `IndexStore`/`EdgeStore`
coordination boundary: SQLite rollback completes before the graph facade is
notified. The RTree backend then reconnects both persistent virtual tables
before publishing its restored live seal. Callers use `IndexStore.transaction()`
rather than invoking those internal rollback hooks directly.

## Later-phase population

P3 populates `list_objects`, `list_object_columns`, `fblocks`, formula
destination edges, and its parse/name/dynamic/inconsistency diagnostic subset.
Verified P4 consumes those edges for graph queries and circular analysis. The
Verified P5 completes the shared catalog, adds cached-error/link/volatile
production, and filters by sheet/severity/code before a 100-row cap. P6
populates staleness and invokes the already typed `I_STALE` constructor; P7
invokes the typed regex-timeout constructor and exposes the bounded MCP/CLI
query.
Until each phase gate closes, the presence of its schema table is not evidence
for a later phase's behavior.
