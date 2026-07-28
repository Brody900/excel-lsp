# Phase 4 dependency graph and trace evidence

Recorded on 2026-07-28 (America/New_York). This report covers the P4 candidate
for M5 dependency queries, bounded traces and paths, and exact circular-reference
verification. Formal R-mech and R-test verdicts are recorded only after they
inspect one frozen staged fingerprint.

## Delivered contracts

- `EdgeStore` provides inclusive point/range overlap through identical R*Tree
  and interval-table interfaces. Keyset pages are bounded to 1–1,000 IDs and
  reject invalid sheet IDs, coordinates, cursors, and limits.
- `DependencyGraph` keeps source nodes at formula-block or singleton-cell
  granularity. Precedents intersect the query with source rectangles and return
  destination references; dependents spatially intersect destination rectangles
  and return source blocks/cells.
- Direct results, bounded breadth-first trees, and dependent-direction shortest
  paths use frozen public `fblock:` and `cell:` symbols. Trees contain
  references, reasons, and emitted node/edge counts—never cell values. A node's
  `child_count` is exact on an untruncated result; at the truncation boundary it
  is a bounded lower-bound witness that at least one child was omitted.
- Trace depth defaults to 2 and is capped at 8. Paths default to three results,
  cap depth at 12, retain distinct parallel `via` edges, and return a normal
  disconnected result when no path exists.
- Trace and path caps are applied during retrieval. Dense ranks freeze the exact
  public hop order independently of edge IDs. Direction-specific 4-D
  `rtree_i32` mirrors index destination/source rectangles plus those ranks, so
  bounded traversal finds successive intersecting ranks without sorting or
  materializing the complete matching fan-out. The interval backend preserves
  exact result parity when RTree is unavailable.
- Relational edges and both complete spatial mirrors are cross-checked when a
  persisted sidecar opens, then each returned edge is checked again when
  queried. Partial rectangles, fractional coordinates, invalid packed cells,
  missing sheets, orphaned blocks, negative ordinals, and unsupported source
  kinds become structured `E_CORRUPT` errors rather than false topology.
- Schema v5 adds dependent/precedent ranks, both ranked mirrors, and a
  trigger-maintained dirty gate plus mutation/clean trust epochs. Every formula
  refresh atomically rebuilds the workbook-wide ranks and mirrors and then
  seals the current mutation epoch. Current-version sidecars are accepted only
  when the state and rank-key tables, their exact columns, all 18
  contributing-table, catalog, and mirror
  triggers match their complete canonical definitions, and no additional
  persistent trigger can mutate graph trust state. Opening a current sidecar
  rederives every canonical public-hop key and proves both relational rank
  columns match the exact dense semantic order, persisted maxima, and the
  `graph_rank_keys` catalog. Exact interval/RTree DDL and active RTree shadow
  integrity are validated too.
  Invalid state normally rebuilds transactionally; corrupt virtual storage that
  cannot be dropped is replaced by a checkpointed same-directory temporary
  database through atomic `os.replace`, preserving `generation + 1`, WAL,
  foreign keys, and the original error on failed replacement. After opening, a
  process-local immutable tuple seals all seven state fields—including both
  rank maxima and revision. A private authorizer-backed graph-write epoch plus
  `data_version` and `schema_version` witness same-handle graph mutations,
  other-handle commits, and DDL; a process-local canonical rank-key snapshot
  binds every selected rank. Statement caching is disabled so repeated prepared
  SQL cannot evade the epoch, while caller authorizers remain chained. The
  public `IndexStore.connection` surface is a narrow capability that wraps
  cursors and their `.connection` property and rejects native-cursor callback
  factories, so native SQLite base descriptors cannot replace the tracker
  through supported public operations. The constant-size live checks remain
  O(1), and queries consume the validated
  maximum rather than rereading mutable metadata.
  All authoritative graph SQL and schema PRAGMAs are bound explicitly to
  `main`. Protected TEMP names/targets are rejected before schema setup,
  construction, and managed mutation; a later TEMP creation invalidates every
  cached surface through the tracked graph epoch or a constant-size
  `temp.schema_version` seal on plain connections. Canonical export explicitly
  binds every projection to `main` and rejects protected TEMP objects. Direct
  writable-schema catalog writes are included in the graph epoch, including
  repeated connection, cursor, and `executemany` forms, while a caller-denied
  graph update is rejected before that epoch can advance.
  Graph rebuilds may nest only inside `IndexStore.transaction()`; caller-owned
  raw SQLite transactions are refused before mutation because Python exposes no
  reliable callback for their later commit or rollback. Managed body, commit,
  rollback-hook, and commit-hook failures always run idempotent no-I/O
  finalizers that restore or publish the process seal and clear transaction
  bookkeeping without attempting rollback after a durable SQLite commit.
  A proven RTree rollback reconnects both virtual tables before resealing the
  epoch, absorbing SQLite's write-shaped but non-mutating authorization replay
  after TEMP-schema rollback; subsequent surfaces must leave the epoch stable.
  Construction validates the sidecar, captures every seal, and warms the
  interval cache inside one `BEGIN IMMEDIATE` snapshot, excluding writers in
  both WAL and DELETE modes. Every complete direct, trace, and path operation
  runs inside one deferred SQLite read snapshot when no caller transaction
  exists. Trust validation is
  the first read, graph-owned snapshots close on every exit, and raw or managed
  caller transactions are never committed or rolled back by a query. Snapshot
  release uses rollback-state inspection and direct-SQL fallback; an
  unreleasable snapshot poisons and conclusively closes the native SQLite
  descriptor. Snapshot `BEGIN` and managed-store `BEGIN IMMEDIATE` acquisition
  are inside those cleanup boundaries because an instrumented native connection
  can take effect and then raise. `IndexStore.close()` uses the same native
  base-descriptor emergency release after bounded virtual retries and proves
  physical closure through the non-virtual base `in_transaction` descriptor
  before marking the store closed. Primary errors, prior causal evidence, and
  cleanup failures remain distinct, identity-unique, and acyclic through query,
  managed-transaction, context-manager, successful snapshot multi-cleanup,
  direct-close aggregation, and post-commit hook/finalizer exits—even when the
  primary is nested inside a later immutable group or multiple members share
  one external causal chain. A shared returning sanitizer also rebuilds a sole
  immutable group at successful-query, close, post-commit hook, and constructor
  release boundaries, preserving group metadata, traceback, cause/context, and
  explicit suppression while removing duplicate memberships and cycles.
  `graph_analysis_version=1` participates in both lifecycle fast paths and
  forces a complete derived rebuild when absent or obsolete.

## Circular-reference algorithm

Circular detection first runs deterministic iterative Tarjan SCC over formula
blocks. Exact verification is restricted to candidate SCCs:

1. Stage 2a streams every formula cell and checks whether an exact translated
   dependency contains that same cell.
2. A singleton block exits cleanly only when all internal dependencies have one
   strict row-major direction, which is a complete acyclicity proof and covers
   running totals.
3. Remaining candidates use at most 64 deterministic corner/quantile seeds and
   bounded exact expansion. A discovered cycle emits one canonical
   `E_CIRCULAR`; exhausted or incomplete coverage emits
   `W_POSSIBLE_CIRCULAR` with the instruction to verify in Excel.

The handoff calls Stage 2b “multi-block,” but a literal singleton stop is
unsound. `A2 = A1+A3` and `A3 = A2+A4` share one R1C1 block, have no direct
self-inclusion, and still contain `A2 -> A3 -> A2`. P4 therefore applies the
bounded fallback to an ambiguous singleton only after the streaming monotonic
proof fails. This conservative correction preserves F09b while detecting the
real homogeneous cycle; it is frozen by an executable regression and an
independent brute-force differential.

Balanced per-sheet rectangle indexes serve both coarse block intersections and
SCC-local exact expansion. The store adapter uses a separate balanced owner
index, so exact source lookup does not scan every formula block. Anchor-cached
ordinary A1 geometry is resolved at the target cell; structured references,
coordinate spills, and composite range hulls are translated and reanalyzed per
cell. Names and 3-D expansions retain the P3 reference context.

## Fixture and golden evidence

P4 adds five deterministic generated workbooks while preserving the locked F01,
F07, and F19 bytes:

| ID | Contract | Locked SHA-256 |
|---|---|---|
| F04 | distinct global and sheet-local named-range edges | `9844b879673deae1455054aca03f20168e10121d61f23e3204235b8ce3574a0c` |
| F05 | native Table1, current-row and column structured refs | `1132fe9ddd6cca3d0a4e4af4a38e4a56d8b1ebbf1f9526e19c6dec053fc1d397` |
| F09a | exact two-cell cycle | `b18a88c6e1c92a25ef0c3de851bd5675278f03d04518f80db6f501416fbf1234` |
| F09b | 50,000-row coarse self-overlap, exact running-total DAG | `0998132a470b1258aa2b2c1a68162f5c7c16b59e56d8848b6110cbd2f8917675` |
| F15 | `SUM(Jan:Mar!B2)` expanded to three sheet edges | `96496cf60b8990e4d97c857cfe00b4593c856e3bb687ec376cdc00fe830a788f` |

`tests/golden/p4-graph-semantics.json` freezes dependent and precedent trees plus
paths for F03/F04/F05/F15/F19, exactly one F09a `E_CIRCULAR` with its
`B2 -> B3 -> B2` path, and no F09b circular diagnostic. Its serialized form is
also held below 4,000 `o200k_base` tokens with pinned `tiktoken==0.13.0`. The
real-fixture integration suite derives every expected public precedent from the
independently hand-authored edge oracle, asserts exact precedent completeness
for each source, and walks each independently expected target back through
dependents on both spatial backends. This proves I12 without enumerating only
the graph output. The same multiset oracles freeze every
F03/F04/F05/F15/F19 semantic edge by source kind, exact fblock or singleton-cell
geometry, destination geometry, and `via`.
F03 contains exactly 13 edges, including `Summary!C7 -> Calc!B2`; omissions,
extras, duplicates, opaque edges, and unexpected cell sources fail independently
of graph output.

## Boundedness and differential evidence

- RTree/fallback point and range queries have exact parity, including whole-column
  point checks at the first and last Excel rows (I14). A deterministic Hypothesis
  property compares random multi-sheet edge sets, boundary rectangles, points,
  and keyset page sizes through both backends against an independent brute scan.
- A SQLite authorizer regression fails if any of the five direct, trace, or path
  query surfaces touch the `cells` table. The real F09b workbook is re-indexed
  with Stage 2a instrumented and Stage 2b's exact graph allocator replaced by a
  hard failure. Stage 2a resolves each of the 50,000 formula cells exactly once;
  the test compares the complete ordered sequence from row 3 through row 50,002,
  so omissions, duplicates, substitutions, and reordering all fail. It proves
  the monotonic DAG, and indexing remains clean without allocating the fallback
  graph (I13).
- A 16-case backend-by-direction matrix tampers persisted source/destination
  rectangles, ranks, missing rows, and orphan rows, then deliberately forges
  both the clean bit and trust epoch to exercise the deeper defenses. Each case
  still returns structured `E_CORRUPT`; mutation checks independently kill
  validators that ignore either direction, ranks, or relational cross-checks.
  A separate bounded-surface matrix deletes or displaces the queried-direction
  row, forges both clean fields, and exercises precedent traces, dependent
  traces, and paths before and after facade construction on both backends.
  Persisted openings reject the complete mirror mismatch; live facades reject
  any changed mutable trust-state field—including either rank maximum and
  revision—against their complete process seal in O(1). Additional
  probes reject `via`-only semantic reranking, coherent dual-mirror rank swaps,
  malformed physical DDL, corrupt RTree node/rowid shadows, and raw transaction
  adoption. Managed commit/rollback/nesting, hook failures before and after
  state publication, idempotent no-I/O finalization, and interval-cache
  restoration are frozen separately. All five public graph query surfaces map
  genuine live SQLite storage failures to `E_CORRUPT` while preserving caller
  misuse (`sqlite3.ProgrammingError`) and non-database exception identity.
- Cross-connection race injection commits a coherent replacement graph after
  entry validation but before retrieval on every public surface and both
  backends. WAL readers return the complete pre-commit version; rollback-journal
  writers block until the reader snapshot releases and then commit. A fresh
  facade observes the new version, proving the test crossed real generations
  without mixing them.
- Real `sqlite3.Connection` subclasses inject failures immediately before and
  after graph `BEGIN`, managed `BEGIN IMMEDIATE`, and physical `close`. The
  checked-in matrices cover all five graph surfaces, both spatial backends,
  WAL and DELETE journals, normal and denied rollback, persistent virtual-close
  failure, close-after-effect, successful and failing context bodies, and body
  errors with an existing explicit cause or implicit context. All cases prove
  transaction release or native descriptor closure, writer unblocking, exact
  primary exception identity, retained evidence without causal cycles, and a
  clean same-database reopen.
- Additional real subclasses override both virtual `close()` and
  `in_transaction` to lie in both directions: a live descriptor raises the
  ordinary closed-database error, while a physically closed descriptor raises
  an unrelated runtime error. Native base-descriptor probes prevent both false
  closure and false openness. Separate graph, managed-transaction, and context
  exit cases reuse one earlier body cause/context inside a multi-error cleanup
  group and traverse every cause, context, and group member by identity; each
  object must appear exactly once.
- `graph_rank_keys` stores one canonical serialized public-hop key for each
  `(direction, rank)` in a `WITHOUT ROWID` composite primary key. Rebuild fills
  it only after both mirrors are complete. Relational, block, sheet, catalog,
  and active RTree/interval mirror triggers invalidate the affected identity or
  direction. Bounded queries perform one indexed selected-rank lookup and
  compare it with the emitted hop. Exact state plus catalog restoration on the
  live connection remains rejected by monotonic process-local tokens, including
  managed same-handle restoration before or after a pending rebuild,
  coordinated relational/catalog restoration from a second connection, and
  trigger removal before mutation. Non-graph writes remain accepted. Valid
  duplicate hops still share one row and emit once. Recursive cleanup
  normalization removes nested primary membership, deduplicates repeated group
  members, assigns a shared external causal identity one owner, and cuts causal
  cycles while retaining every distinct exception object once.
- A 262-case mechanics regression matrix covers both backends and all five graph
  surfaces for coherent managed restoration with and without a pending rebuild;
  repeated SQL through connection, cursor, and `executemany`; caller authorizer
  chaining, denied graph writes with unchanged authority, writable-schema
  catalog mutation, protected TEMP shadows, accepted non-graph writes and
  rollback; constructor
  barriers in WAL and DELETE; `BEGIN IMMEDIATE` failures before and after native
  effect in both journal modes; caller-owned transactions; writer-first
  ordering; nested primaries; shared external causal chains; sole-group
  rebuilding at every standalone boundary; and constructor cleanup before and
  after native acquisition effect. It also freezes plain-connection TEMP DDL,
  protected-shadow canonical export, complete primary-graph exclusion, and
  exception-group subclass derivation with custom state. Independent TEMP
  variants enumerate all 33 protected names on both backends and isolate name,
  target, case-folding, object-type, and both RTree shadow-prefix predicates.
  Capability-boundary variants prove the public connection, returned cursors,
  cursor `.connection`, and context-manager value never become native
  `sqlite3.Connection` objects; custom cursor/row factories are rejected while
  caller authorizer observation and denial semantics remain intact. Constructor
  cleanup variants inject failures before and after writer-lock acquisition and
  before and after virtual close takes effect, requiring the exact initialization
  primary, unique acyclic cleanup evidence, native closure, and immediate writer
  reacquisition. Direct storage regressions additionally cover capability
  allocation immediately after open; early and final connection-configuration
  failures; schema rollback before and after native effect; replacement-schema
  failure; post-commit replacement-close failure; and temporary-descriptor close
  failure before and after effect. Each proves exact primary identity, physical
  closure, writer availability, original-database preservation where replacement
  has not occurred, and zero rebuild artifacts.
  A separate both-backend `SQLITE_IGNORE` regression deletes a real edge while
  suppressing trigger updates, proves the private epoch still advances, and
  requires all five previously sealed graph surfaces to reject the stale state;
  the denial regression continues to prove rejected writes advance nothing.
  Compound graph, commit, and constructor cleanup tests preserve a primary's
  pre-existing explicit cause or visible context exactly once.
- Permanent both-backend kill matrices prove the exact catalog columns, both
  checks, every `NOT NULL`, composite-primary-key composition/order, and
  `WITHOUT ROWID` identity; reject nine independent DDL mutations; corrupt
  missing, wrong, and extra catalog rows in both directions after exact
  seven-field restoration; and individually
  delete or mutate all 18 trigger definitions. Each current-sidecar corruption
  is rejected directly and reconstructed with monotonic generation and exact
  canonical post-rebuild content. An independent frozen manifest and SQL oracle
  cover all 18 triggers. Functional INSERT/UPDATE/DELETE matrices assert the
  exact post-operation catalog for every base/catalog family and both active
  mirror directions.
- SQLite `EXPLAIN QUERY PLAN` assertions exercise the actual ranked spatial
  lookup and reject temporary B-tree sorting. With a progress callback every
  100 SQLite VM operations, two-node traces over 1,000/10,000/50,000 irrelevant
  edges require 3/5/8 callbacks for dependents and 4/8/11 for precedents. The
  capped prefix remains identical to the complete semantic order after more
  than 1,000 edge IDs are reversed.
- A committed 6,000-decoy circular regression inspects at most eight coarse leaf
  candidates and exactly two SCC-local candidates per exact dependency.
- Iterative Tarjan matches a brute SCC oracle across 300 deterministic random
  graphs and handles 1,505 nodes without recursion.
- Path tests cover cycles, disconnected graphs, multiple shortest parents,
  parallel reasons, layered combinatorial truncation, and depth limits.

These timing figures are engineering work-bound evidence, not the frozen P8
release benchmark. P8 records public index/build latency, tokens, accuracy, and
incremental-series charts under §9.

## Fresh verification

Current pre-freeze results:

| Check | Result |
|---|---|
| Graph, EdgeStore, lifecycle, circular, property, real-fixture, catalog, and mechanics slice | 1,202 passed in 119.48s |
| Commit/constructor/cleanup/schema mechanics plus exact catalog integrity | 332 passed in 12.48s |
| Real F09b Stage 2a/2b boundedness guard | 1 passed in 17.76s |
| Full repository suite with branch coverage | 1,938 passed in 365.82s; 90.05% core coverage |
| Fixture generation + parser oracle | 21 passed in 17.36s |
| Ruff check / format | clean |
| Pyright | 0 errors, 0 warnings |
| Deterministic fixture generator | 14 workbooks regenerated; locked hashes reproduced |
| Lock check / package build | clean; sdist and wheel built |

Exact commands used for the final refresh (the result table is updated only
from these commands):

```powershell
uv run pytest tests/unit/test_graph.py tests/unit/test_circular.py tests/unit/test_index_store.py tests/unit/test_index_lifecycle.py tests/unit/test_p4_graph_fixtures.py tests/unit/test_p4_rank_catalog_integrity.py tests/unit/test_p4_mechanics_regressions.py tests/property/test_edge_store_properties.py -q
uv run pytest tests/unit/test_p4_mechanics_regressions.py tests/unit/test_p4_rank_catalog_integrity.py -q
uv run pytest tests/unit/test_p4_graph_fixtures.py::test_i13_real_f09b_never_allocates_the_stage_2b_exact_graph -q
uv run pytest --cov=excel_lsp.core --cov-branch --cov-report=term --cov-fail-under=85
uv run pytest tests/unit/test_fixture_generation.py tests/oracle/test_oracle.py -q
uv run ruff check .
uv run ruff format --check .
uv run pyright
uv run python tests/fixtures/generate.py
uv lock --check
uv build
```

## Formal phase gate

The approved candidate used base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`fdfa3d0e00faba1844d9636e650b63ee8a395a4c`, and cached-diff blob
`67fb916469f7a109b1f584907c0128f1d7164f3d` across 34 files with 18,748
insertions and 349 deletions. The single combined follow-up reviewer verified
that fingerprint unchanged at entry and exit and returned clean verdicts:

- global verdict #49 / R-mech #26: `APPROVE`
- global verdict #50 / R-test #21: `APPROVE`

The prior combined candidate's global #47/#48 `REVISE` verdicts remain charged
in `PLAN.md`. Non-formal preflights never substituted for a formal verdict.
