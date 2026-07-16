# Phase 2 regions, symbols, and workbook-map evidence

Recorded on 2026-07-15 and updated on 2026-07-16
(America/Los_Angeles). This phase implements M2: sparse region detection,
headers and column profiles, the frozen public symbol IDs, and the bounded
workbook map used for semantic navigation.

## Delivered contracts

- Coordinate-ordered sparse row runs with blank-row/blank-column gap tolerance
  defaulting to 1 and validated in the authoritative range 0–8; no dense
  worksheet grid is constructed.
- Native Excel ListObjects take precedence, reproduce their declared ranges
  exactly, provide authoritative column headers, and act as hard barriers to
  heuristic regions.
- Pairwise-disjoint heuristic rectangles, deterministic coordinate ordinals,
  formula-without-cache occupancy, and height-independent merged-range
  structural spans.
- Header inference across at most three rows with merged-text synthesis and
  safe bold/fill style evidence.
- Exact nonnull counts, type-tagged distinct estimates capped at 1,000, and
  deterministic dtype samples capped at 200 values per column (50 at reduced
  rate for a sheet above 2,000,000 non-empty cells).
- Frozen sheet, region, column, name, formula-block, and cell symbol formatters,
  including Unicode NFKC header normalization and `#2`, `#3`, ... collision
  suffixes.
- Atomic persistence of regions, columns, and `W_LARGE_SHEET` with cells; index
  schema version 2 invalidates pre-region sidecars, and `gap_tol` participates
  in freshness so a configuration change is never mistaken for a hash no-op.
  Omitted freshness calls preserve a valid same-workbook tolerance; only an
  explicit override changes it.
- Deterministic map degradation, an 8,000-character hard bound, hidden and
  veryHidden visibility plus exact omission counts by visibility, `hasVBA`,
  capped names/links, exact diagnostic counts, and no body-cell values.
- Seven fixed index reads use response-cap-derived ceilings of 200 sheets, 80
  round-robin regions, 16 base columns per selected region plus 512 bounded
  extras, 20 names, and 10 sanitized external-link labels while retaining exact
  totals and remainder counts.

## Header-scoring decision

The accepted candidate maximizes this score, with ties favoring fewer header
rows and a threshold of 0.55:

| Feature | Weight |
|---|---:|
| textual header coverage | 0.30 |
| header/body type contrast | 0.25 |
| unique normalized synthesized headers | 0.20 |
| bold/fill shift from representative body style | 0.20 |
| nonblank header coverage | 0.05 |

If no candidate reaches the threshold, the region has zero header rows and a
column-letter fallback such as `Column B`. For merged headers, each merge
anchor contributes once per output column, producing names such as
`Revenue / Q1` rather than repeating the anchor text.

## Invariant evidence

| Invariant | Evidence |
|---|---|
| I4 every non-empty cell belongs to at most one region | Unit cases cover adversarial intersecting component bounds and table-wrapping shapes. Non-vacuous Hypothesis properties assert pairwise-disjoint output and exactly one containing region for every planted cell in supported random grids. |
| I5 ListObject ranges are exact | Table validation and properties compare the exact planted table rectangle/name multiset, assert authoritative headers/totals exclusion, and prove no heuristic rectangle intersects a table. F01 persists `SalesTable` as exactly `A1:D6`. |
| I6 ordinals are deterministic | Seeds are sorted by natural coordinates after table/heuristic normalization. Properties vary stream chunking and table/merge metadata order while asserting identical regions, columns, and symbol IDs. |

## Sparse performance regressions

The region engine combines a coordinate-ordered row sweep, deterministic
spatial lookup for ListObject barriers, a dynamic interval/sparse-grid
rectangle index, and queue-driven bounding-box closure. Full-height vertical
barriers are batched only for the consecutive ordered prefix that commutes;
partial-height barriers retain the original above/below/left/right BSP order.
Merged headers are assigned through one sparse rectangle index and queried
through at-most-three-row interval views rather than rescanning every merge for
every candidate region and cell.

The corrected 50-by-50 isolated-grid instrumentation produces 2,500 regions
with 39,096 total candidate operations: 23,100 public `Rect.intersects` checks,
15,996 sparse exact checks, and zero interval-report candidates. The former
all-pairs paths performed 6,247,500 rectangle comparisons on the same input.
The committed regression counts every current candidate source and holds the
combined work below 80,000.

Additional committed adversarial gates cover:

- 500 to 1,000 full-height ListObjects over 12 rows: copied-run work grows
  exactly from 12,024 to 24,024, replacing a preflight reproduction that grew
  from 1,515,012 to 6,030,012 copied items;
- a 512-rectangle cascading staircase, a 512-rectangle dense overlap, and an
  empty-center case where only expanded component bounds intersect;
- 512 same-row merged-header regions with merge-metadata iteration bounded
  linearly;
- 400 table-partition examples against the former ordered BSP behavior,
  including reversed table metadata, and 80 bounding-closure examples against
  a brute fixed-point reference.

Finite orchestrator smoke probes on this machine completed full analysis of
1,000 full-height ListObjects over 12,012 sparse cells in 0.482 seconds, the
2,000-rectangle staircase in 0.361 seconds, and 4,000 merged-header regions in
1.509 seconds. These timings are recorded as smoke evidence only; the committed
acceptance gates use deterministic work counts and exact differentials.

The third mechanics remediation replaces the rect-key-balanced BVH with a
spatially partitioned tree whose iterator still emits exact intersections in
deterministic rect-key order. It also replaces one-run-per-merged-row expansion
with exactly one raw span per anchored merge. Ordinary cell runs are coalesced
once, mixed runs and spans are clipped algebraically by the same table BSP, and
a row-expiry plus column-interval sweep computes exact sparse components. The
ordinary no-merge path retains its existing row-run fast path.

Fresh non-verdict adversarial preflight produced the following evidence:

- reviewer-style partial-height decoys with 800, 1,600, and 3,200 total tables
  required 21,137, 45,489, and 97,393 exact checks, respectively; output hashes
  and counts were identical with reversed table metadata, and doubling
  exponents were 1.106 and 1.098 rather than quadratic;
- an independent interleaved-decoy series at 400, 800, and 1,600 total tables
  required 4,089, 8,977, and 19,553 checks;
- the maximum-height merged range `A1:A1048576` reaches mixed partitioning as
  one ordinary anchor run plus one span and returns exactly one region; its
  primitive count is independent of height;
- `gap_tol=8` is accepted, `gap_tol=9` is rejected before index creation, and
  maximum-tolerance same-column work grew from 8,955 to 17,955 to 35,955 unions
  for 1,000, 2,000, and 4,000 consecutive rows;
- 96 then 192 tall side-by-side merges with equal counts of unrelated sparse
  cells and table metadata created exactly 96 then 192 spans, 288 then 576
  mixed input primitives, and 288 then 576 output regions;
- one independent random differential compared raw spans with explicit
  per-row expansion across 30,000 geometries, tolerances 0–3, and both metadata
  orders: 240,000 comparisons completed with zero mismatches. A separate audit
  added 229,376 exhaustive small-grid comparisons and 100,000 randomized
  comparisons, also with zero mismatches.

The committed unit and property regressions retain bounded versions of each
gate. The larger finite probes are additional pre-review evidence, not release
benchmarks.

## Deterministic fixture evidence

P2 adds six generator-built fixtures while preserving the exact P1 bytes for
F01 and F07:

| Fixture | P2 evidence |
|---|---|
| F02 | Three islands at default tolerance; `gap_tol=0` deterministically splits the one-blank-row island and advances generation. |
| F03 | Three exact ListObject regions for Inputs, Calc, and Summary; compact reference map. |
| F12 | Two-row merged headers synthesize `Region`, `Revenue / Q1` through `Revenue / Q3`, and `Units / Actual` / `Units / Target`. |
| F13 | Exact `int`, `date`, `float`, `float`, `str`, `bool`, `mixed` column sequence; dates are inferred from raw serials plus number formats. |
| F14 | Empty sheets remain indexed with no regions; `B2` and `X100` are deterministic singleton regions. |
| F20 | 40 sheets, 12 ranked islands, 300 names (60 of each defined-name kind), hidden and veryHidden sheets, eight regions plus `{"more":4}`, and 20 names plus `namesMore:280`. |

Every one of the eight currently generated fixtures has byte-identical repeated
generation and exact production-parser versus pinned-openpyxl oracle equality.

## Map budgets

The committed normalized counts in
`benchmarks/results/map-budgets.json` use compact UTF-8 JSON semantics and
`tiktoken==0.13.0` with `o200k_base`:

| Fixture | Serialized characters | Tokens | Required cap |
|---|---:|---:|---:|
| F03 | 1,041 | 342 | 1,500 tokens and 8,000 characters |
| F20 | 4,575 | 1,595 | 8,000 characters |

The exact normalized maps are committed as
`tests/golden/f03-workbook-map.json` and
`tests/golden/f20-workbook-map.json`. Runtime timestamps are the only field
replaced by a placeholder before comparison. An additional test builds a map
for every currently generated fixture and enforces the 8,000-character bound.

## Repository claim discipline

The P2 README skeleton follows the required persuasion-document order but marks
unimplemented P3-P9 behavior and all unmeasured performance claims explicitly.
`docs/evidence/readme-claims-to-artifacts.md` maps every README claim to its
producing phase and exact future proof. Public setup is Codex-first; generic MCP
JSON is labeled as compatibility rather than Codex's native TOML format.

## Verification and review

Fresh orchestrator verification on 2026-07-16:

| Check | Result |
|---|---|
| `uv lock --check` | passed; 70 packages resolved |
| `uv run python tests/fixtures/generate.py` | passed; F01, F02, F03, F07, F12, F13, F14, and F20 regenerated |
| `uv run ruff check .` | passed |
| `uv run ruff format --check .` | passed; 42 files already formatted |
| `uv run pyright` | passed; 0 errors, warnings, or information messages |
| `uv run pytest --cov=excel_lsp.core --cov-report=term-missing --cov-fail-under=85` | 254 passed; 90.02% branch coverage |
| `uv build` | built the `excel_lsp-0.1.0` sdist and wheel |
| Markdown local-link audit | 28 Markdown files checked; every local target resolved |
| tracked-junk audit | 82 tracked paths checked; no tracked sidecars, generated workbooks, caches, OS junk, or build output |
| `git diff --check` | passed |

Formal P2 review #1 returned `REVISE` in both implementation domains. R-mech
found that an omitted map read reset a stored `gap_tol=0`, extreme degradation
could hide non-visible sheets behind a visibility-neutral count, and source
loading used unbounded/N+1 region and column queries. R-test found that public
external-link labels could expose URL credentials, query strings, and fragments.

The remediation preserves stored analysis configuration, adds exact
`sheetListMoreByVis` counts, uses seven scale-independent bounded SQL reads, and
sanitizes URL-aware labels without exposing authority or secret-bearing suffixes.
The tiny and 121-sheet scaling indexes execute the same seven reads; the larger
case contains 1,452 regions, 58,080 columns, 500 names, and 15 links while the
in-memory projection stays at 121 sheets, 80 regions, 1,792 columns, 20 names,
and 10 links. A separate 205-sheet index proves the sheet projection stops at
its cap-derived ceiling of 200 while prioritizing non-visible sheets. A
bounded-versus-unbounded differential scan was identical for caps from 350
through 8,000 characters; the first excluded full renders measured 8,312
characters for 201 minimal sheets, 8,918 for 81 minimal regions, and 9,665 for
513 minimal columns. Focused remediation and golden tests pass.

Formal R-test review #2 independently returned a clean `APPROVE`. The single
early R-repo invocation returned `REVISE`: it found weakened S5 wording, no
executable-artifact mapping for `excel-lsp bench`, and an underspecified
raw-results index. The candidate includes changes for those findings, but no
later verdict has approved the P2 repository gate.

Fresh formal re-reviews #2 returned `REVISE` in both remaining domains. R-mech
found quadratic scaling paths in per-ListObject run partitioning and zone
repartitioning, repeated global overlap-closure sweeps, and per-region/per-cell
merged-header scans. The remediated engine now uses ordered barrier lookup with
linear full-height batching, live component bounds with full-batch absorption,
and sparse merged-header interval views. Non-verdict preflight caught and
closed two residual defects before another formal invocation: copied run views
were still quadratic for side-by-side tables, and batching all full-height
barriers could cross an intervening partial-height barrier. The final
500-to-1,000 copy-work gate is exactly linear, the supplied semantic
counterexample is committed, and an additional 30,000-case two-order
orchestrator differential found no table-partition mismatch.
R-repo found that six current-state documents still described the early review
with a single-invocation rule and misstated the authorized retry protocol. The
current documentation and its contract test now use the authorized fresh
re-review protocol consistently and keep every earlier invocation charged.

Fresh formal reviews #3 produced different outcomes. R-repo returned `APPROVE`
with one minor: `docs/architecture.md` still implied that R-test and all three
P2 approvals were pending. This current-state update records that R-test and
R-repo have approved and that only fresh R-mech approval remains.

R-mech returned `REVISE` with three majors. Rect-key BVH queries traversed
quadratic decoy branches: 400 tables against 203 zones performed 82,931 checks,
while 800 tables against 403 zones performed 326,622. A 100,000-row anchored
merge expanded into 100,001 row runs instead of remaining a lazy rectangle or
span. Unrestricted `gap_tol` bridging performed 499,500 union calls for 1,000
same-column runs and 1,999,000 calls for 2,000. The candidate now implements
the spatial iterator, one raw span per anchored merge, sparse component sweep, and the
authoritative maximum tolerance of 8 described above. Focused tests, the full
suite, static checks, directed cases, and the independent raw-span
differentials are green. This is remediation evidence, not a mechanics
approval.

P2 remains gate-pending solely on a fresh R-mech approval; read-only preflights
do not consume the ledger.

Formal R-mech review #4 returned `REVISE`. Its independent tests confirmed that
the spatial iterator, one-span-per-merge representation, gap-work cap, maps,
static checks, and full 255-test suite were otherwise green, but found a
remaining global barrier path. The mixed-geometry BSP applied each table row to
every span in one coarse zone before proximity components existed. For
10, 20, 40, and 80 tall merges plus equal unrelated one-cell tables, heuristic
output grew to 220, 840, 3,280, and 12,960 regions; direct span constructions
grew through 76,800 at 160. A horizontally disjoint `M2` table could even split
`A1:A20` into three heuristic regions. The required correction is
component-local barrier partitioning: compute sparse proximity plus
bounding-box closure first, then apply only tables intersecting each component
bound and recompute children without rejoining across the barrier. The review
also requested runtime rejection of boolean and non-integer `gap_tol` values to
prevent persisted-configuration drift.

This verdict raises the ledger to 15 of 30 and consumes the last contingency.
P2 remains gate-pending, and the next fresh mechanics invocation is the final
available review under the authorized overall ceiling.

## Component-first remediation after review #4

The remediated merge-bearing path now computes primitive proximity and
bounding-box closure before considering any ListObject. Two primitives within
the configured row and column gaps connect only when a literal cell-to-cell
witness rectangle is table-free. Horizontal and vertical gaps use projected
interval coverage; diagonal gaps use the nearest boundary rectangle. Tables are
then applied in deterministic rectangle order only to components whose bounds
they intersect. Every directional child is recomputed independently, so
siblings cannot reconnect across the barrier. The ordinary no-merge row-run
path is unchanged.

The proximity sweep evaluates a static witness edge when its later primitive
arrives. It keeps every processed root in an all-root bounds index for closure
and only row-near roots in a second index for witness work. Each root owns
immutable power-of-two member-index blocks. The blocks meld binomially, discard
historical roots, and use exact spatial traversal, which keeps live membership
linear and prevents already-internal or axis-only decoys from reappearing as
primitive candidates. Boolean and non-integer `gap_tol` values now fail before
sidecar creation.

Non-verdict preflight deliberately counted internal traversal and construction
work, not only returned regions:

- a literal witness oracle completed 201,632 exhaustive and 100,000 seeded
  random multi-table cases with zero mismatch;
- an independent fixed-order component/BSP oracle completed 20,001 geometries
  in both normal and reversed table metadata order (40,002 optimized runs) with
  zero mismatch, invariant, mutation, or order failure;
- the output-sensitive connected wraparound family produced exactly 1,721 and
  6,641 regions and 5,740 and 22,680 child fragments at `n=40` and `n=80`;
  total counted work was 53,361 and 225,365, a 4.2234× increase for the
  inherently quadratic semantic output;
- the strongest two-component same-root-decoy family produced two regions
  while total work grew 16,281 → 35,373 → 76,301 for
  `n=160 → 320 → 640`; its formerly hidden axis candidates grew only
  2.0013× and exact spatial checks 2.1893× at the final doubling;
- unrelated in-bounds tall merges at `n=64,128,256` required exactly
  256, 512, and 1,024 rectangle checks, zero span splits, and returned
  192, 384, and 768 regions;
- a maximum-height `A1:A1048576` anchored merge remained one span and one exact
  region with roughly 22 KiB measured peak memory, independent of height;
- named diagonal blockers, the `D1`/`D2` directional cases, the expired-root
  closure case, strict runtime tolerance types, and lifecycle refusal all
  passed.

The independent post-fix snapshot used
`regions.py` SHA-256
`58a8f85dae2fde434015eb9c25f61d75e3f704f033f8b2dccac1da04c15ff0ef`.
The full repository suite reported 278 passing tests and 89.96% branch
coverage for `excel_lsp.core`; repository-wide Ruff, format, Pyright, and diff
checks were clean. These results were exhaustive pre-review evidence, not a
mechanics approval. At that checkpoint, P2 remained gate-pending on the final
fresh stateless R-mech invocation.

## Gate result

Fresh stateless P2 mechanics review #5 (`R-mech` invocation #9, overall verdict
#16) returned `APPROVE` with no critical, major, or minor findings. The reviewer
independently reran 82 focused tests and the complete 278-test suite, Ruff,
formatting, Pyright, the lock check, staged-diff checks, 101,990 literal witness
comparisons, the prior global-fragmentation family, balanced-axis and
same-root scaling at 640/1,280, 600/1,200 spatial decoys, 6,000 exact
spatial-order queries over 1,800 rectangles, and maximum-height merge probes.
The staged `regions.py` hash matched the documented digest and the unstaged
worktree remained empty throughout.

With its earlier clean R-test and user-authorized early R-repo approvals, this
verdict closes P2. The review ledger now records 16 of 30 invocations used and
14 pooled invocations remaining, exactly matching the minimum P3–P9 completion
path while retaining three R-repo invocations for P9.
