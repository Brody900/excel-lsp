# Excel LSP agent decision log

This file is append-only. Every non-trivial project decision records the decision, alternatives, and rationale.

## 2026-07-15 P0 — Orchestration and release target

Decision: execute HANDOFF v1.1 phases P0–P9 in order on `main`, with milestone commits, frozen review budgets, and a public v0.1.0 release as the target.

Alternatives considered: one large implementation pass; a reduced prototype; delaying repository/release work until after implementation.

Rationale: the handoff makes phase gating, review accounting, evidence, and release artifacts part of correctness. Incremental milestones keep a large greenfield build reviewable and recoverable.

## 2026-07-15 P0 — Semantic retrieval fallback

Decision: do not use the available Serena session for this repository; use `rg`, targeted file reads, and native language tooling until a correctly scoped session exists.

Alternatives considered: query the existing Serena session; proceed without checking its scope.

Rationale: Serena reported that it was activated for a different local project, not this workspace. Querying another project would produce misleading retrieval results and violates the workspace-scoping rule.

## 2026-07-15 P0 — Primary parser architecture

Decision: use a single-pass, namespace-tolerant lxml OOXML parser as the production read path; keep openpyxl only for deterministic fixture creation, formula utilities, and the oracle.

Alternatives considered: openpyxl read-only dual loads; openpyxl normal-mode loading; a hybrid parser with openpyxl as the primary source.

Rationale: streaming lxml supports the S1 performance target, exposes tables/merges/styles directly, shares package infrastructure with the surgical editor, and avoids relying on unverified openpyxl read-only behavior. The oracle retains an independent correctness check.

## 2026-07-15 P0 — Codex execution and repository instructions

Decision: treat every Claude reference in the historical handoff as Codex, keep `CLAUDE.md` only for the frozen Definition of Done, and add `AGENTS.md` as the active Codex instruction file. Do not invoke Claude or Anthropic services.

Alternatives considered: delete `CLAUDE.md`; mechanically substitute command names and leave incompatible flags; follow the Anthropic-specific examples unchanged.

Rationale: the user explicitly required Codex, current Codex CLI semantics differ, and keeping the compatibility file satisfies the handoff without making it the active authority. A repo-local `HANDOFF.md` preserves the full specification and records this substitution at the top.

## 2026-07-15 P0 — Machine lane and spatial backend

Decision: use the Windows COM live-test lane and SQLite R*Tree backend. Develop in an explicit uv-managed Python environment supporting Python 3.11–3.13.

Alternatives considered: macOS AppleScript; skip live Excel; use the indexed interval-table fallback; inherit the agent process's Python environment.

Rationale: Windows 11 x64 and Excel 16.0 build 19530 were confirmed, a hidden COM launch/quit succeeded, Python 3.11–3.14 and uv 0.9.30 are installed, and SQLite 3.47.1 successfully created and queried an RTree virtual table. The fallback remains part of the portable code contract but is not the active machine path.

## 2026-07-15 P0 — Headless Codex isolation

Decision: run benchmark agents through `codex exec --ephemeral` with verified Codex-native MCP/profile controls; retain the authenticated user configuration and isolate project rules/tools explicitly.

Alternatives considered: the handoff's `claude -p --bare` command; `codex exec --ignore-user-config`; omitting LLM evaluations.

Rationale: a normal-config read-only JSON ping returned `PONG`, while `--ignore-user-config` repeatedly failed authentication with HTTP 401 even though `codex login status` was healthy. The benchmark harness must preserve authentication while reproducing the intended isolation, tool allowlist, turn guard, and usage capture.

## 2026-07-15 P0 — Distribution name and deterministic pins

Decision: retain the preferred distribution name `excel-lsp`; pin `openpyxl==3.1.5` and `tiktoken==0.13.0` exactly.

Alternatives considered: the authorized fallback `excel-lsp-mcp`; older dependency releases; unbounded minimum-version declarations for the deterministic dependencies.

Rationale: both the PyPI JSON endpoint and project page for `excel-lsp` returned 404 on 2026-07-15, while the selected versions were the current stable PyPI releases and support Python 3.11. Exact pins make fixture XML and token-count evidence reproducible across the CI matrix.

## 2026-07-15 P0 — VBA source handling

Decision: commit only the extracted `xl/vbaProject.bin`, ignore the local `f16_source.xlsm`, and record both hashes and macro provenance in `tests/fixtures/README.md`.

Alternatives considered: commit the source workbook; synthesize a VBA project; omit F16 until the live phase.

Rationale: the source workbook contains personal/local metadata and is not a release artifact. `olevba` verified that its only active macro is `Stamp`, a hidden read-only Excel COM run produced `Z1 = 42`, and the source SHA-256 remained unchanged before and after execution. The extracted 13,312-byte blob matches SHA-256 `be05aafbb31d2de0ffd686c9cae71b97a2596132ed4443d9d656558d7089ccb1`.

## 2026-07-15 P0 — Handoff count reconciliation

Decision: interpret “21 generated fixtures” as 21 fixture scenarios and exercise every emitted workbook, including all F06 sizes and both F09 variants. Treat the audit-cost callout as the fifth rendered chart while keeping the accuracy table as Markdown.

Alternatives considered: generate only 21 physical workbooks; count the accuracy table as an image; weaken the literal Definition of Done.

Rationale: the fixture table names 21 scenarios but necessarily emits additional variant files, and the chart section specifies four conventional plots plus a numeric callout while the Definition of Done requires five charts. This interpretation satisfies every concrete artifact requirement without discarding variants or turning the accuracy table into an inaccessible image.

## 2026-07-15 P0 — Public commit identity

Decision: use the authenticated GitHub account name and its GitHub-provided no-reply address in this repository's local Git configuration.

Alternatives considered: inherit the machine-global personal email; invent a project address; delay identity configuration until push.

Rationale: the user requested a public repository, so commit attribution should work without publishing an unrelated private email. The change is repository-local and does not alter global Git configuration.

## 2026-07-15 P0 — Type-checking boundary

Decision: use Pyright basic mode for the repository and strict mode for `src/excel_lsp/core` from the first production module.

Alternatives considered: basic mode everywhere; strict mode across server, CLI, tests, and benchmarks immediately.

Rationale: the core is the reusable correctness boundary and benefits from strict typing, while framework glue and test code can remain in basic mode unless later evidence justifies expanding strictness. The scoped configuration is cheap at greenfield stage and matches the handoff's preferred upgrade.

## 2026-07-15 P0 — Scaffold self-check

Decision: accept the Phase 0 scaffold and close its self-check gate.

Alternatives considered: begin parser work before a clean baseline; defer build and coverage checks to Phase 1.

Rationale: fresh orchestrator runs passed lock verification, locked all-extras sync, Ruff lint and format, Pyright, four tests with 100% scaffold core coverage, fixture-generator execution, CLI execution, and sdist/wheel build. The branch is `main`, the local authoring workbook is ignored, and the public artifact tree contains no generated indexes or credentials. Full command evidence is recorded in `docs/evidence/p0-recon.md`.

## 2026-07-15 P1 — Typed parser and index boundary

Decision: make the lxml package parser emit immutable typed records through a callback, persist the normalized stream in an internal `cells` table, and expose one shared JSON-scalar normalization function for all downstream public values.

Alternatives considered: return worksheet-sized lists from the parser; let each tool normalize values independently; defer the typed contracts until region detection.

Rationale: callback emission preserves the single-pass streaming constraint, the internal cell table lets later phases derive regions and formula blocks without reparsing, and one normalization boundary prevents date, boolean, and error values from drifting between SQLite, goldens, samples, and tool responses.

## 2026-07-15 P1 — Stable lifecycle and canonical comparison

Decision: store selected package hashes separately, place default indexes beside the workbook, name centrally configured indexes `<stem>.<first8-sha256-absolute-path>.xlsp.db`, and define canonical exports entirely in natural workbook keys rather than SQLite surrogate ids.

Alternatives considered: key central indexes by filename alone; compare raw auto-increment rows; rebuild every sheet whenever the file stat changes.

Rationale: hashed absolute paths avoid collisions between same-named workbooks, natural-key exports make invariant I2 testable across independent databases, and selected-part hashes permit true per-sheet refresh while workbook/shared-string/style changes trigger the required broader invalidation. Generation advances once per mutation and untouched stat matches remain no-ops.

## 2026-07-15 P1 — Opaque edge compatibility

Decision: retain nullable destination sheet and coordinate columns in the frozen `edges` schema, and insert spatial rows only for edges with real destination rectangles.

Alternatives considered: represent opaque/external destinations with zero coordinates; make every destination coordinate non-null until graph work begins.

Rationale: later reference extraction must represent external and dynamic references with no destination rectangle. Sentinel coordinates would violate worksheet bounds and make spatial queries return false matches, so Phase 1 preserves the schema contract before Phase 3 consumes it.

## 2026-07-15 P1 — Deterministic fixtures and openpyxl read-only VERIFY

Decision: generate F01 and F07 deterministically, inject known cached formula values, convert F07 fill-down formulas into genuine shared groups, and keep the openpyxl oracle exception list empty for cell streams.

Alternatives considered: compare only formula text without cached values; anticipate an openpyxl shared-formula skip; use read-only worksheet metadata for tables and merges.

Rationale: under pinned `openpyxl==3.1.5`, `ReadOnlyWorksheet` correctly expands F07 shared followers at C3, C11, C14, and C21, so no cell-stream skip is justified. Read-only `.tables` and `.merged_cells` each raise `AttributeError`; a normal-mode control sees `FormulaBlocksTable` and `E1:F1`, confirming that these are read-only API deficiencies rather than missing fixture content. Both generated F01 and F07 have exact production-parser versus dual-load openpyxl cell-stream equality. Reproduce the informational probe with `uv run python tests/oracle/probe_openpyxl.py` or `uv run python -m tests.oracle.probe_openpyxl`; detailed expectations live in `tests/oracle/skiplist.md`.

## 2026-07-15 P1 — Adversarial parser hardening

Decision: retain integral parser values as Python integers, coerce only values outside SQLite int64 to finite REAL storage, require and enforce shared-formula spans, parse typed `calcPr` metadata, and represent What-If Data Tables as non-textual typed formula spans with their first/second input-cell attributes.

Alternatives considered: reject large integral numerics; store them as text; synthesize a fictional `TABLE()` formula string; accept missing shared spans; defer calculation metadata.

Rationale: Excel numerics use a floating-point domain even when their lexical representation is integral, while SQLite cannot bind arbitrary Python integers. Boundary-only coercion preserves normal integers and avoids crashes. Shared and Data Table master spans are required to attribute members safely. Data Table formula elements intentionally contain metadata rather than formula text, so a typed opaque record is honest and future-editor-safe. The parser also now records all relevant `calcPr` settings instead of silently assigning defaults.

## 2026-07-15 P1 — Complete incremental dependencies

Decision: include each worksheet relationship part and its referenced ListObject parts in selected package hashes, associate them with the owning sheet, and preserve incoming cross-sheet edges when only a destination sheet is replaced.

Alternatives considered: hash only worksheet XML; force a full reindex on every whole-file hash change; delete every edge touching a reindexed sheet.

Rationale: a table rename or range change can modify only `xl/tables/table*.xml`; treating the new file stat as fresh without reindexing would permanently retain stale region metadata and break I5. Conversely, an unchanged source formula's edge remains valid when its destination sheet's cell content changes, so deleting destination-only edges would make incremental and full canonical exports diverge. Workbook-structure changes still clear and rebuild the complete catalog and edge set.

## 2026-07-15 P1 — Concurrent schema initialization

Decision: serialize schema creation and migration with `BEGIN IMMEDIATE`, re-check schema state after acquiring the SQLite write lock, execute DDL without `executescript`'s implicit commit, and use bounded busy/locked retries for initialization pragmas and lock acquisition.

Alternatives considered: rely only on `busy_timeout`; use process-local thread locks; make all DDL `IF NOT EXISTS`; accept first-open races as a deployment limitation.

Rationale: multiple MCP/server processes can cold-open the same derived index. A process-local lock would not protect that case, and idempotent individual DDL still permits partial interleaving. SQLite's own cross-process write lock gives one atomic initializer or migrator; later connections re-check and reuse the completed schema. An eight-thread, 30-database adversarial probe improved from intermittent `database is locked`/duplicate-table failures to zero failed runs.

## 2026-07-15 P1 — First adversarial gate cycle

Decision: treat both first P1 reviewer invocations as REVISE and fix every reported major finding before spending the second R-mech and R-test invocations.

Alternatives considered: classify the findings as future-phase work; pass the phase on green aggregate coverage; weaken the parser or concurrency contracts.

Rationale: R-test identified missing `calcPr`, locked-open retry, and parse-time torn-save retry coverage. R-mech reproduced valid Data Table rejection, table-part invalidation, incoming-edge deletion, and concurrent cold-schema races. Additional live review probes exposed large-integer SQLite overflow, unconstrained shared followers, incomplete F07 cache assertions, and incomplete built-in date-format coverage. These were root-cause defects or meaningful blind spots in the Phase 1 foundation, so all were closed with focused regressions rather than deferred.

## 2026-07-15 P1 — Hash-stable freshness and physical-id-free evidence

Decision: treat an equal whole-package hash as a semantic no-op even when source timestamp metadata changes or is malformed, update only the persisted stat bookkeeping transactionally, and map every spatial sheet coordinate to its natural sheet name in canonical exports.

Alternatives considered: advance generation whenever the file timestamp changes; leave malformed stat metadata unrepaired; expose R*Tree or interval sheet ids as canonical evidence because production sheet ordering is deterministic.

Rationale: timestamp-only changes do not alter workbook meaning and must not invalidate generation-bound cursors. The equal-hash path therefore re-stats before committing, retries once on a concurrent save, preserves generation, and repairs the fast-path metadata. Canonical exports are evidence across independently built databases, so even normally deterministic physical ids are implementation details; cross-database tests now vary sheet and edge ids for both spatial backends and still produce identical exports while preserving edge-to-rectangle association.

## 2026-07-15 P1 — Review interruption accounting

Decision: charge the frozen review ledger only for protocol-complete reviewer invocations that return the required `VERDICT` artifact, while recording two pre-verdict process interruptions separately.

Alternatives considered: count a process as a review invocation when it starts even if it never returns a verdict; omit the interruptions from project history.

Rationale: one read-only reviewer was stopped after an unbounded delegated probe and a later replacement was lost when the task automatically restarted; neither produced findings, an approval, or any gate artifact. Counting process failures as completed reviews would make the mandatory remaining P2–P7 mechanics gates arithmetically impossible within the frozen ten-verdict budget. Both interruptions are disclosed here; no worktree reset was needed because neither process edited files.

## 2026-07-15 P1 — Phase gate accepted

Decision: close Phase 1 after the fourth verdict-bearing R-mech invocation returned a finding-free APPROVE and the second R-test invocation returned APPROVE with its documentation-only minor corrected.

Alternatives considered: begin region work before an approved mechanics gate; spend another R-test verdict immediately on an already-corrected evidence-table mismatch.

Rationale: the final mechanics reviewer independently inspected the full P1 scope and observed 153 focused tests passing with no findings. Fresh orchestrator evidence is stronger and broader: 160 tests pass at 85.75% branch coverage, Ruff lint and format checks pass, Pyright reports zero findings, deterministic fixtures and the pinned openpyxl probe reproduce, the lock is current, the package builds, and `git diff --check` is clean. The R-test minor did not affect code or test adequacy and is corrected in `PLAN.md` and `docs/evidence/p1-foundation.md`.

## 2026-07-15 P2 — Sparse region engine and header weights

Decision: detect heuristic islands from coordinate-ordered sparse runs with a default one-blank-row/column tolerance, make ListObjects hard spatial barriers, and score up to three candidate header rows with weights 0.30 textual coverage, 0.25 type contrast, 0.20 normalized uniqueness, 0.20 style shift, and 0.05 nonblank coverage at a 0.55 threshold.

Alternatives considered: construct a dense grid; infer regions from worksheet `<dimension>`; allow heuristic rectangles to overlap table blanks; use header strings alone; require bold formatting.

Rationale: sparse runs scale with workbook content rather than worksheet bounds, while exact tables satisfy I5 and remove ambiguity. The mixed score handles plain exports, styled reports, and multi-row merged headers without making formatting mandatory. Pairwise rectangle normalization and non-vacuous properties give a stronger form of I4, and coordinate sorting makes I6 independent of union-find roots, relationship order, or SQLite ids.

## 2026-07-15 P2 — Column identity and profiling

Decision: normalize column headers with Unicode NFKC, case folding, punctuation/whitespace collapse, Unicode alphanumeric preservation, and deterministic column-letter fallback; suffix normalization collisions left-to-right with `#2`, `#3`, and so on. Profile exact nonnull counts and type-tagged distinct values up to 1,000 while sampling at most 200 dtype values.

Alternatives considered: expose raw headers directly in symbol ids; ASCII-only slugs; hash headers; treat booleans as integers or error strings as text; compute unbounded exact distinct counts.

Rationale: the frozen `col:{sheet}:{region}:{normHeader}[#k]` scheme needs readable, repeatable components that cannot inject `:` or `#`. Type tags keep `TRUE` distinct from `1`; int-plus-float promotes to float; stored date strings remain dates through `value_type`. Saturation bounds memory without pretending to provide exact high-cardinality statistics.

## 2026-07-15 P2 — Analysis freshness and map degradation

Decision: bump the derived index schema to version 2, persist a region-analysis version and `gap_tol`, and require both values on the stat/hash fast paths. Build the map in workbook order, rank displayed regions by area without renumbering their coordinate IDs, cap regions/names/external links at 8/20/10, and apply deterministic secondary degradation until compact JSON fits 8,000 characters.

Alternatives considered: preserve P1 sidecars with empty region tables; treat a tolerance change as a no-op against equal workbook bytes; paginate the map; renumber regions in display order; rely on P7 to truncate an oversized P2 map.

Rationale: analysis configuration is part of semantic index identity even though it is not workbook content. The map is the cheap orientation response, so pagination would defeat its purpose and physical or display-order IDs would be unstable. The F03 normalized map measures 342 `o200k_base` tokens, while F20 measures 4,575 characters with all 40 sheets and both non-visible states surfaced.

## 2026-07-15 P2 — Codex-first public skeleton

Decision: create the README in the frozen §10.2 order, use verified Codex MCP syntax and native TOML as the primary setup path, label generic `.mcp.json` as compatibility, and map every current or future public claim to a named phase and committed evidence path.

Alternatives considered: retain historical Claude commands from the handoff; omit the README until benchmarks exist; fill comparison and benchmark sections with projected results; defer the claims matrix to release week.

Rationale: the user explicitly defined historical Claude references as Codex, and the local CLI confirms `codex mcp add excel-lsp -- uvx excel-lsp serve`. Honest placeholders make the repository understandable now without selling unimplemented behavior. The early matrix lets the one required P2 repository review catch missing proof plans before P9, when retrofitting live captures and raw benchmark links would be expensive.

## 2026-07-16 P2 — Sparse overlap indexing

Decision: replace all-pairs region-bound normalization and disjoint validation with a self-ordering row sweep, fixed-width column buckets, exact candidate checks, and repeated DSU bounding-box closure; reuse the same overlap enumerator for table and merge validation.

Alternatives considered: retain the simple quadratic loops; rely on wall-clock thresholds; add a general-purpose interval-tree dependency; stop after optimizing only the final disjoint assertion.

Rationale: a valid 2,500-cell isolated grid reproduced 6,247,500 rectangle-intersection calls and multi-second analysis. The sparse sweep reduces that probe to 46,200 calls while preserving the former closure semantics. The committed regression gates on candidate work rather than machine timing, and a 100-example Hypothesis property compares the indexed pair set with brute force, including wide rectangles and reversed input order.

## 2026-07-16 P2 — Pre-review proof hardening

Decision: finish P2 preflight before spending its one remaining mechanics verdict, one remaining test verdict, and one-shot repository verdict; strengthen exact serialized goldens, tokenizer metadata, degradation branches, fixture caches, transactional rollback, persisted warnings, current-claim links, and future-claim path accounting.

Alternatives considered: invoke formal reviewers as soon as aggregate tests passed; treat parsed JSON equality as an exact golden; leave unexecuted degradation branches and future documentation paths for P9.

Rationale: the frozen review arithmetic leaves exactly one mechanics and one test verdict for each remaining implementation gate, while P2 permits exactly one repository invocation. Preflight found one real quadratic path and several evidence gaps that ordinary green coverage did not expose. Closing them first produced 216 passing tests at 88.69% branch coverage, exact F03/F20 serialization and budget checks, complete F03 cache expectations, and an auditable README skeleton without consuming a formal verdict.

## 2026-07-16 P2 — Formal review remediation

Decision: charge the first formal P2 R-mech and R-test verdicts as `REVISE`, close every reported major before re-review, preserve stored region-analysis configuration on omitted freshness calls, project the map through fixed SQL limits with visibility-aware omission totals, and sanitize public external-link labels through URL-aware basename extraction.

Alternatives considered: treat the map as an administrative read that may reset `gap_tol`; retain a single visibility-neutral `sheetListMore`; keep application-level per-sheet/per-region queries because the serialized result is small; expose the raw external target or hostname when no safe path basename exists; reclassify either formal verdict as a preflight.

Rationale: map reads must not mutate semantic configuration, and a bounded response does not justify unbounded source loading. The remediated loader performs seven reads on both tiny and 121-sheet indexes and uses response-cap proofs to bound its projection at 200 sheets, 80 round-robin regions, 16 base columns per selected region plus 512 extras, 20 names, and 10 links while retaining exact totals for degradation. A 205-sheet regression crosses the sheet ceiling without changing the query count. Visibility counts keep hidden and veryHidden state discoverable even when identities no longer fit. URL labels discard authority, query, fragment, parameters, unsupported schemes, malformed targets, encoded delimiters, and non-text or noncanonical metadata entries while the raw valid index metadata remains available for later reference resolution. Fresh verification now reports 235 passing tests at 88.56% branch coverage; no verdict or review spend was hidden.

## 2026-07-16 P2 — Test approval and repository-plan correction

Decision: count the second formal R-test invocation as a clean `APPROVE`, count the required early R-repo invocation as `REVISE`, and correct every repository finding without weakening a frozen success criterion: restore S5's at-least-10× naive-baseline threshold, add an exact future CLI test/evidence route for `excel-lsp bench`, and enumerate the raw benchmark filenames and acceptance gates.

Alternatives considered: summarize S5 as any positive reduction; assume the benchmark runner scripts prove the advertised CLI command; leave raw-result filenames distributed only across the claims matrix; treat the R-repo findings as future P8/P9 concerns; omit the `REVISE` because the early repository review was intended as a single invocation.

Rationale: the claims-to-artifacts plan is specifically meant to make later evidence mechanically auditable. A weaker paraphrase could falsely certify S5, scripts alone do not prove CLI wiring, and a raw-results link should tell contributors exactly what must exist and what makes it acceptable. The formal R-test reviewer independently reran the P2 security, golden, property, fixture, and coverage evidence and returned no findings. The repository verdict remains charged and visible even though its findings are documentation-only.

## 2026-07-16 P2 — Review-governance pause

Decision: do not spend the formal P2 mechanics re-review or a repository re-review until the user authorizes a transparent amendment to the frozen review allocation; continue all non-verdict remediation, testing, evidence, and staging work meanwhile.

Alternatives considered: combine later phase verdicts in one reviewer process; treat an interim statement as a second phase approval without charging it; skip the repository re-review after `REVISE`; silently transfer a repository slot to mechanics; proceed into P3 before P2 closes.

Rationale: five mechanics slots remain, but P2 reapproval and the five P3-P7 mechanics gates require six fresh invocations. The early P2 repository invocation also returned `REVISE`, so the general protocol requires a new invocation even though the budget section says to spend exactly one at P2. The minimum honest completion path uses 26 of the 30 total reviews if unused domain slots are pooled and a P2 repository retry is allowed, leaving four contingencies while preserving every gate and at least three P9 repository reviews. Until authorized, the original domain caps and one-shot wording remain authoritative.

## 2026-07-16 P2 — Review-governance amendment authorized

Decision: activate the user's explicit authorization to pool otherwise-unused review-domain slots under the unchanged 30-invocation ceiling, permit the required second P2 R-repo invocation after its first `REVISE`, preserve every phase gate and fresh stateless reviewer requirement, retain all prior verdict charges, and reserve at least three R-repo invocations for P9.

Alternatives considered: keep the project paused; waive the P2 repository re-review; combine multiple later gates into one verdict; raise the overall budget above 30; alter or erase prior review accounting.

Rationale: the amendment resolves the mechanical and repository retry arithmetic without weakening the review protocol or Definition of Done. Ten verdicts have been used, the minimum complete path still totals 26, and four pooled contingency invocations remain available while the P9 repository reserve stays protected.

## 2026-07-16 P2 — Second formal re-reviews returned REVISE

Decision: charge formal P2 R-mech re-review #2 and R-repo re-review #2 as `REVISE`, keep P2 gate-pending, update the ledger to R-mech 6 used, R-test 4 used, and R-repo 2 used, and require fresh stateless re-reviews after focused remediation. Twelve of 30 verdicts are now used, 18 pooled verdicts remain, the minimum completion path is 28, two contingency verdicts remain, and at least three R-repo invocations stay reserved for P9.

Alternatives considered: treat the re-reviews as preflights; preserve stale exactly-one or one-shot repository wording; count the first repository `REVISE` as satisfying the gate; continue into P3; describe the findings as fixed before focused verification and new verdicts exist.

Rationale: R-mech found three major scaling defects. ListObject hard-barrier processing rescanned every table for adjacent runs and repartitioned all accumulated zones once per table, with a 12,000-cell probe growing from 6.1 seconds at 1,000 tables to 23.6 seconds at 2,000; remediation requires indexed table candidates and a many-ListObject work-bound regression. Cascading rectangle overlaps could trigger repeated global sort-and-sweep closure passes, with a 2,000-rectangle staircase taking 17.8 seconds; remediation requires queue-driven component expansion backed by spatial candidate lookup and a cascading-closure regression. Merged-header inference scanned the complete merge list per candidate region and then linearly searched it per header coordinate, with 4,000 sparse merged-header regions taking 3.7 seconds; remediation requires indexed region assignment or direct coordinate-to-anchor lookup plus a many-merge scaling regression. R-repo found that `PLAN.md`, the claims plan, the evidence index, the P2 evidence draft, index internals, and the changelog still described the R-repo gate as exactly-one or one-shot and retry authorization as pending. Current-state documentation must instead state that the first required early invocation remains charged as `REVISE` and the user-authorized fresh re-review protocol continues until approval or the documented overall exhaustion policy.

## 2026-07-16 P2 — Indexed region scaling remediation

Decision: replace the three reviewer-identified quadratic paths with rect-key-ordered ListObject barrier lookup and consecutive full-height batching, live dynamic interval/grid component indexing with full-batch bounding-box closure, and one sparse merged-header index with at-most-three-row interval views. Gate the changes with operation counts and exact differentials rather than machine-time thresholds.

Alternatives considered: retain the first indexed draft after its original focused tests passed; loosen timing thresholds; batch every full-height barrier regardless of intervening ordered barriers; count only `Rect.intersects`; use dense merge-coordinate maps; proceed directly to another formal verdict.

Rationale: non-verdict preflight reproduced a residual side-by-side-table path at 6,030,012 copied runs and 19.245 seconds for 1,000 barriers, and found that the isolated-grid test did not observe the new private intersection path. The final 500-to-1,000 barrier regression copies exactly 12,024 then 24,024 runs, while corrected isolated-grid instrumentation totals 39,096 candidate operations versus the original 6,247,500 comparisons. A broader differential then caught one semantic error in the batching optimization: a later full-height barrier was being moved across an earlier partial-height barrier. Batching is now limited to the consecutive rect-key prefix, the exact counterexample is committed, the property runs 400 examples in both metadata orders, and an additional 30,000-case two-order preflight found no mismatch. Independent smoke probes completed full analysis of 1,000 full-height tables over 12,012 cells in 0.482 seconds, the 2,000-rectangle cascade in 0.361 seconds, and 4,000 merged-header regions in 1.509 seconds. Fresh formal mechanics and repository verdicts remain required.

## 2026-07-16 P2 — Third formal reviews split

Decision: charge formal P2 R-mech review #3 as `REVISE` and R-repo review #3 as `APPROVE` with one minor, keep P2 gate-pending solely on a fresh R-mech approval, and update the ledger to R-mech 7 used, R-test 4 used, and R-repo 3 used. Fourteen of 30 verdicts are now used, 16 pooled verdicts remain, the minimum completion path is 29, one contingency verdict remains, and at least three R-repo invocations stay reserved for P9.

Alternatives considered: treat the repository minor as a failed repository gate; leave architecture implying that R-test and all three approvals are pending; treat the mechanics probes as optional optimization; proceed into P3; describe unimplemented mechanics remediation as complete.

Rationale: R-repo approved the early repository gate but identified one current-state minor in `docs/architecture.md`: its P2 wording still implied that R-test and all three required approvals were pending. The current architecture must instead state that R-test and R-repo have approved and that only fresh R-mech approval remains. R-mech found three major scaling defects. Rect-key BVH decoy traversal performed 82,931 checks for 400 tables and 203 zones, then 326,622 for 800 tables and 403 zones; remediation requires adaptive exact spatial candidates with deterministic ordering plus a decoy-work gate. Anchored merges expanded one run per covered row, so a 100,000-row merge produced 100,001 runs; remediation requires lazy rectangles or spans plus a tall-merge regression. Unrestricted `gap_tol` bridging performed 499,500 union calls for 1,000 same-column runs and 1,999,000 for 2,000; remediation requires component/indexed sweeping or an authoritative validated small maximum plus an adversarial regression. No mechanics remediation or approval is recorded by this decision.

## 2026-07-16 P2 — Spatial, tall-merge, and tolerance remediation

Decision: replace the decoy-prone rect-key BVH with a spatially partitioned tree that preserves exact rect-key output order; replace per-row anchored-merge expansion with compressed row-band spans and algebraic ListObject clipping; bound `gap_tol` to the authoritative range 0–8 and reject larger values before index creation; retain the existing bounding-box closure after a sparse row-expiry and column-interval component sweep.

Alternatives considered: materialize and sort every exact table candidate on every query; keep only merge endpoint sentinels; connect raw merge rectangles before table partitioning; preserve unrestricted tolerance and optimize arbitrary-distance row joins; accept timing-only regressions; begin P3 before a fresh P2 mechanics approval.

Rationale: the spatial tree removes the reviewer-supplied decoy traversal while retaining the BSP's deterministic table order. Row-band events at populated rows, merge boundaries, and table boundaries reproduce the former per-row horizontal coalescing without work proportional to merge height. A maximum of 8 preserves the documented configurable behavior around the default of 1 while limiting active-row comparisons to a small constant; lifecycle validation prevents partial sidecars. Fresh preflight measured 21,137, 45,489, and 97,393 checks for 800, 1,600, and 3,200 reviewer-style tables, with exact reverse-order output. `A1:B1048576` remained four runs, one span, one zone, and zero unions at 7,576 bytes peak. Maximum-tolerance union work scaled 8,955 → 17,955 → 35,955 for 1,000 → 2,000 → 4,000 rows. An independent row-expanded oracle completed 60,018 production detector runs with zero mismatches (seed 20260716; rolling hash `61ff18b8ee9f595cf1f7e66c5e1beb04bd377020776f1ab68737b0469ad51b55`). These are non-verdict remediation results; P2 remains gated on one fresh stateless R-mech approval.

## 2026-07-16 P2 — Raw-span correction after non-verdict audit

Decision: supersede the compressed row-band implementation before formal review. Coalesce ordinary cell runs once, retain exactly one raw span for each anchored merged range, partition the mixed primitives through the existing ListObject BSP, and use the sparse component sweep only for sheets that contain anchored merges. Preserve the original ordinary-sheet component fast path.

Alternatives considered: keep the green row-band candidate because it passed the reviewer-supplied cases; cap the number of merges or row events; cache complete active-merge signatures; accept an output-sensitive `O(active merges × row events)` path; spend the fresh formal verdict before another preflight.

Rationale: a read-only audit proved that 40, 80, and 160 tall disjoint merges plus equal unrelated row events processed 3,360, 13,120, and 51,840 band inputs and emitted 3,280, 12,960, and 51,520 intermediate spans—approximately 4× work for each 2× input increase. The raw-span replacement creates 96 then 192 spans and 288 then 576 total mixed primitives for the corresponding doubled committed gate. Root independently compared it with explicit per-row expansion for 30,000 geometries, tolerances 0–3, and both metadata orders (240,000 comparisons) with zero mismatches. A second audit completed 229,376 exhaustive small-grid and 100,000 randomized comparisons with zero mismatches. The spatial decoy fix and tolerance cap remain unchanged. P2 is still gate-pending; this correction spent no formal review invocation.

## 2026-07-16 P2 — Fourth mechanics review found global barrier fragmentation

Decision: charge formal P2 R-mech invocation #4 as `REVISE`, keep P2 gate-pending, and update the ledger to R-mech 8 used, R-test 4 used, and R-repo 3 used. Fifteen of 30 verdicts are now used. The minimum completion path consumes all 30 available invocations, so the next fresh P2 mechanics review is the final available invocation and no pooled contingency remains.

Alternatives considered: treat the green raw-span equivalence oracle as proof of the legacy global BSP; classify quadratic empty-region output as harmless metadata; waive the final mechanics approval; proceed into P3; edit during the formal review; omit the runtime tolerance-type minor because static typing says `int`.

Rationale: the reviewer confirmed one major. The global mixed run/span BSP still sliced every tall merge at every unrelated table row before component closure. With n disjoint tall anchored merges on the left, n one-cell tables at distinct rows on the right, and adjacent non-table cells keeping one coarse zone, n=10, 20, 40, and 80 produced 220, 840, 3,280, and 12,960 heuristic regions in 0.022, 0.085, 0.344, and 1.431 seconds. Direct span constructions reached 300, 1,200, 4,800, 19,200, and 76,800 for n=10, 20, 40, 80, and 160. Even `A1:A20` fragmented into `A1`, `A2`, and `A3:A20` because of a horizontally disjoint `M2` table. Remediation must compute sparse proximity plus bounding-box components before applying a table BSP, then split only components whose bounds intersect that table and keep disjoint spans atomic. The reviewer also reported a minor: `bool` and non-integer `gap_tol` values must be rejected at runtime to prevent persisted configuration drift. Focused 105 tests, the full 255-test suite, Ruff, formatting, Pyright, and 50,000 exact spatial-order queries were otherwise green; the reviewer made no worktree changes.

## 2026-07-16 P2 — Component-first regions and final-verdict preflight

Decision: replace the global mixed-geometry barrier pass with a table-aware later-primitive component sweep, all-root bounding-box closure, and fixed-order component-local BSP. Define primitive proximity by the existence of a table-free minimal cell-to-cell witness; use projected interval coverage for horizontal and vertical corridors and the nearest boundary rectangle for diagonal corridors. Retain root membership in immutable power-of-two spatial-index blocks that meld binomially, keep a separate row-active root index for witness discovery, query immutable blocks through exact spatial traversal, and reject boolean or non-integer `gap_tol` values before lifecycle side effects.

Alternatives considered: split all spans globally and accept output proportional to unrelated table rows; apply table barriers before components; use a whole-corridor `intersects_any` shortcut that incorrectly blocks parallel witness rows; count only calls to the final witness predicate; cache one member index for every historical component; rescan flat retained-member lists; use a single global primitive index and filter same-root members after enumeration; accept balanced-axis candidates because wall time was still modest; spend the final formal verdict before exhaustive non-verdict work.

Rationale: component-first semantics preserves an unrelated `A1:A20` merge atomically while still producing exact directional `D1` and `D2` children around a real barrier. Several deliberately adversarial preflights prevented another premature formal invocation. Recomputing every primitive pair after each child made the connected-wrap family super-output-linear. An all-root implementation then revisited blocked roots and retained quadratic historical member indexes; eager flat membership also copied quadratically. The later-primitive invariant removed old-old rechecks, and root-local binomial blocks removed historical and same-root work. A final audit found that the generic balanced-axis intersection path still enumerated 89,651 then 343,155 internal candidates for two output components at `n=320,640`; immutable member blocks now use the exact spatial iterator instead. Post-fix, that family performs 19,131 then 41,883 exact spatial checks, and its total work ratio is 2.157.

Independent preflight on `regions.py` SHA-256 `58a8f85dae2fde434015eb9c25f61d75e3f704f033f8b2dccac1da04c15ff0ef` completed 201,632 exhaustive plus 100,000 seeded random witness comparisons and 20,001 full component/BSP geometries in both table metadata orders with zero mismatch. The canonical wrap family returned the exact 1,721 and 6,641 regions and 5,740 and 22,680 fragments; fully counted work grew 53,361 to 225,365, or 4.2234× for inherently quadratic output. Unrelated tall merges doubled exact work, maximum-height merge memory remained near 22 KiB, strict tolerance and lifecycle checks passed, the full suite reported 278 passing tests and 89.96% branch coverage for `excel_lsp.core`, and Ruff, formatting, Pyright, and diff checks were clean. No review invocation was consumed by this work. The candidate remains P2 gate-pending until the final available fresh stateless mechanics verdict.

## 2026-07-16 P2 — Final mechanics approval and phase closure

Decision: charge fresh stateless P2 R-mech review #5 as a clean `APPROVE`, close every P2 checklist item, update the ledger to R-mech 9 used, R-test 4 used, and R-repo 3 used, and proceed to the P2 milestone commit before starting P3. Sixteen of 30 verdicts are now used, 14 pooled invocations remain, the minimum P3–P9 path uses all 14, and at least three R-repo invocations remain reserved for P9.

Alternatives considered: treat the exhaustive preflight as the required verdict; reuse an earlier reviewer; combine P2 approval with P3 review; begin P3 before recording the gate; leave active-phase language in public status documents; alter or erase the four charged P2 mechanics revisions.

Rationale: the final reviewer inspected only the frozen staged candidate from P1 HEAD `8651386f08b729343f4334549b099c3339ac5177` and returned no critical, major, or minor findings. Fresh independent commands produced 82 focused and 278 full passing tests, clean Ruff/format/Pyright/lock/diff checks, 101,990 literal witness comparisons, exact prior-fragmentation and maximum-height behavior, a 2.1605 fully counted same-root scaling ratio at 640→1,280, 33,115→71,019 checks for 600→1,200 spatial decoys, and 6,000 exact spatial-order queries over 1,800 rectangles. The reviewer verified the staged `regions.py` SHA-256 and confirmed no worktree mutation. P2's R-mech, R-test, and early R-repo gates are therefore all approved.

## 2026-07-16 P3 — Formula semantics and modern shared translation

Decision: build reference extraction on the pinned openpyxl tokenizer behind a
source-faithful compatibility wrapper; normalize the frozen modern-function
catalog before callable/name lookup; implement lexical LET/LAMBDA bindings; and
replace production shared-formula expansion with Excel LSP's tokenizer-backed
A1 translator.

Alternatives considered: treat every non-legacy function as an unknown name;
bind a LET declaration inside its own value; keep openpyxl Translator as the
production shared-group expander; reject spill, implicit-intersection, or
structured-escape shared groups.

Rationale: stored `_xlfn.`/`_xlws.` formulas and modern syntax are normal Excel
OOXML, not parse failures. Desktop Excel confirmed that a LET binding is not
visible inside its own value. Adversarial parser probes proved openpyxl
Translator leaves `@A2` and an `A2:INDEX(...)` endpoint unchanged and raises on
spill/escaped-structured formulas. The replacement matches openpyxl exactly on
a committed 300-example supported-grammar property while dedicated regressions
cover modern constructs that openpyxl cannot translate safely.

## 2026-07-16 P3 — Exact blocks with semantic structured contexts

Decision: normalize every formula cell independently to R1C1, grow
column-major equal runs and merge exact adjacent columns, then extract
references once per block. For bare source-dependent structured operands,
retain the one block and reclassify only that direct operand over homogeneous
ListObject-context rectangles.

Alternatives considered: one edge interpretation from only the block anchor;
split one formula block per table or per cell; re-tokenize every cell; propagate
structured references found indirectly inside defined-name bodies as if they
were direct operands.

Rationale: adjacent tables can contain identical `=[@Input]` formulas with one
R1C1 signature but different source columns. Direct-operand provenance plus
semantic tiling preserves I9/I11 and exact edge coverage without fragmenting
the formula identity. Three deterministic Hypothesis differentials compare
against brute per-cell classification and exact context partitions. Replacing
an active-context rescan with a touched-column row-event sweep reduced the
1,024-context adversarial probe from about 5.16 seconds to 0.012 seconds.

## 2026-07-16 P3 — Dynamic reference expression purity

Decision: determine `INDEX` reference context from complete matched token spans,
including purely parenthesized endpoints and openpyxl's synthetic `:(` group
tokens. Propagate nested `INDEX`, `OFFSET`, `INDIRECT`, and reference-returning
`CHOOSE` results only when the complete branch remains a pure reference
expression.

Alternatives considered: emit when INDEX closes immediately before a colon;
emit as soon as INDEX opens after a colon; treat every INDEX call as
reference-valued; omit nested reference-returning CHOOSE branches.

Rationale: suffix-only state produced order-dependent false positives for
`(1+INDEX(...)):A5` and `A1:(INDEX(...)+1)`, while an unconditional nested rule
misclassified `INDEX({1,2},1)`. Matched whole-expression spans distinguish
those scalar cases from arbitrary parenthesized range endpoints. A 1,440-case
independent endpoint differential matched the resulting model before formal
review.

## 2026-07-16 P3 — Correctness-first normalization and honest S1 status

Decision: discard both cross-cell Translator reuse and a custom fused R1C1
lexer; keep exact per-cell tokenizer semantics for the P3 gate. Record the
current 50,000-formula and complete 50,000-by-10 timings without claiming S1,
and carry the profiled parser/region optimization path to the P8 benchmark
gate.

Alternatives considered: retain a one-row or maximum-row Translator probe;
ship the sub-second fused lexer after only the simple fill-down benchmark;
omit the known cold-index miss from P3 evidence.

Rationale: Translator reuse collapsed distinct `=@A2` and
`=A2:INDEX(A:A,2)` formulas, and boundary probes did not constitute a general
proof. Fuzzing the fused lexer found name tokens such as `C$a4` that it split
and partially rewrote. The safe exact baseline is slower but reviewable:
median R1C1 normalization is 4.154 seconds for 50,000 formulas, while block
construction is 0.036 seconds. The complete current cold path is 37.880
seconds; profiling identifies row-end OOXML parsing and a conservative exact
dense-region accumulator as the material route to the P8 target.

## 2026-07-16 P3 — Typed callable reference flow

Decision: represent reference identity and callable identity separately through
LET/LAMBDA inference. Preserve lexical closures, first-class named LAMBDAs,
higher-order callable arguments and results, and conservative callable
alternatives returned by `CHOOSE` or `IF`. When a computed callable result is a
range endpoint, emit a visible `opaque:<callable>` edge and `I_DYNAMIC_REF`;
retain nested INDEX attribution through simple reference-preserving wrappers.

Alternatives considered: treat every local or defined LAMBDA result as a
scalar; treat every callable result as a reference; recognize only direct
inline LAMBDA calls; discard argument identity at higher-order boundaries; emit
only the innermost dynamic function for arbitrarily deep compositions.

Rationale: Excel permits callables to be passed, selected, returned, named, and
invoked after grouping. Scalar-only inference silently lost real computed
endpoints, while unconditional promotion produced false dynamic edges for
arithmetic and scalar-returning controls. Typed bindings keep those cases
separate. The committed matrix includes `Apply`, `Pick`, `Make`, inline and
lexical closures, callable selectors, scalar controls, both endpoint
directions, and transparent grouping.

## 2026-07-16 P3 — Unified composite reference grammar

Decision: use one quote-, bracket-, and 3-D-aware A1 endpoint grammar for
classification, shared translation, and R1C1 normalization. Fold exact static
colon and representable intersection operators through transparent
parentheses and `@` groups. Compute compatible bounds across A1, names, whole
axes, and structured operands; preserve context-free column/name ambiguity;
and emit explicit opacity for computed or geometrically unrepresentable cases.

Alternatives considered: keep three partially overlapping endpoint parsers;
split every colon at the first textual delimiter; reduce all intersections to
bounding boxes; emit independent endpoint edges without the combined range;
silently drop an operator when its result cannot fit the frozen geometry
model.

Rationale: tokenizer edge forms include synthetic `:(` function tokens,
triple-colon whole-axis/name syntax, escaped structured headers, quoted 3-D
spans, and independent `@` or `#` modifiers. Late non-verdict fuzzing exposed
grouped whole-column, name-to-structured, and extra-parenthesized named or
lexical callable endpoints that earlier focused cases did not cover. Exact
compatible geometry plus visible conservative fallback fixes those losses.
Permanent regressions now cover both directions and nested wrappers; the
context-free translation path remains deliberately non-committal where only a
workbook-defined name can disambiguate the spelling.

## 2026-07-16 P3 — No-verdict final hardening cycle

Decision: delay the formal P3 freeze after non-verdict adversarial passes found
four additional root-cause defects. Retain composite range endpoint geometries
through block extrusion; release all selected sheets' stale ListObject aliases
inside the atomic refresh transaction; flatten, identity-deduplicate, and cap
callable alternatives at 32 with conservative overflow; normalize `:@(` as a
real colon plus grouped implicit intersection; and require exact directional
translation before extending a noncanonical R1C1 block bucket.

Alternatives considered: spend the formal verdicts on the earlier green
candidate; bound only formula length; accept anchor-exact mixed ranges; insert
replacement tables sequentially and special-case sheet order; make every
computed callable opaque without type flow; use only a symmetric block key; or
force every spill formula into a singleton.

Rationale: the mixed endpoint `E3:E10 = SUM(A<row>:SalesTable[Qty])` persisted
only `A2:B5` because the moving endpoint was inside the fixed table range at
the anchor; component-wise extrusion now yields `A2:B10`, with a deterministic
150-example two-dimensional/boundary differential. Moving a table alias from a
later-order sheet to an earlier one collided with its stale row; batch
pre-release permits the move while preserving rollback and real collision
errors. Duplicate callable choices grew from 2.743 seconds at depth 17 to
11.238 seconds at depth 19; normalization reduces the 14-level reproducer from
about 3.68 seconds to about 0.0065 seconds and retains possible reference
results beyond the cap. Both-sided `@(A1):@(B5)` had produced an empty callable
and `W_PARSE`; a 144-case depth/name/structured matrix is now exact.

R1C1 alone also collapsed `A1:A1` with `A2`, nested degenerate composite
endpoints, and directional lowercase/absolute column spellings. Explicit range
arity is now retained, shared translation preserves Excel's absolute-column
case rules, and proposed vertical/horizontal merges verify the top-left formula
at the candidate coordinate. Canonical uppercase formulas retain a fast path.
Coordinate spill operands follow the frozen §5.4 exception: block translation
holds their anchors verbatim while still translating and checking every other
reference in the formula. The final 50,000-formula-column medians are 4.547
seconds normalization, 0.195 seconds guarded block construction, and 0.369
seconds inconsistency analysis. No formal review invocation was spent in this
cycle.

## 2026-07-27 P3 — Computed-name endpoint identity and bounded expansion

Decision: distinguish concrete range names from constants and formula/LAMBDA
names at the classification boundary. Preserve every body precedent, but never
use those precedents as exact result geometry for colon or whitespace
intersection. Carry result-reference and dynamic-function provenance through
LET and alias chains, emitting `opaque:<FN>` plus `I_DYNAMIC_REF` for dynamic
results and `opaque:ref` for scalar or bare-callable endpoints. Treat whitespace
intersection as a full reference context. Memoize resolved name expansion only
inside one task-local analysis scope, keyed by context, scope, anchor, token,
spill/function mode, and complete recursion stack.

Alternatives considered: accept generic block opacity after preventing the
false hull; mark every name use opaque; infer exact output geometry from the
first body precedent; keep colon-specific dynamic detection; cache on the
`ReferenceContext` across workbook cells; or rely only on the depth-32 guard.

Rationale: `Pick = INDEX($A:$A,1)` had been persisted as an exact
`A1:B1048576` hull in `Pick:B5`, while `Scalar = ABS($A$1)` fabricated `A1:B5`.
The first containment fix retained precedents but lost `opaque:INDEX` and
`I_DYNAMIC_REF`; direct or named-LAMBDA INDEX results on whitespace
intersection also became silently nonopaque. Parenthesized constants could
drop the operator entirely. Finally, valid `_F31 -> ... -> _F0 -> Fixed` alias
chains expanded exponentially because folding, inference, and extraction each
reanalyzed the same bodies. Typed result metadata closes the semantic gaps,
and the scoped memo reduces the valid 32-name reproducer to under 0.01 seconds
without sharing state across formulas or weakening recursive-name containment.
Direct, reverse, grouped, block-extrusion, persistence, structured-negative,
and alias-depth regressions are committed. No formal review invocation was
spent in this hardening cycle.

## 2026-07-27 P3 — Occurrence-complete intersections and execution-local memoization

Decision: scan every unfolded whitespace intersection in encounter order and
retain occurrence-specific dynamic markers while emitting at most the required
generic conservative fallback. Deduplicate dynamic diagnostics semantically by
code and function identity when a nested computed-name body is reused as a
colon endpoint. Scope the defined-name memo to the actual execution owner,
represented by retained `Thread` and asyncio `Task` objects, and retain the
exact `ReferenceContext` object in an identity key.

Alternatives considered: stop after the first intersection; globally dedupe
same-label dynamic events; dedupe diagnostics only by full message text; allow
copied contexts to share a mutable cache; use copy-on-write mappings; identify
threads only by recyclable numeric ids; or disable memoization under
concurrency.

Rationale: formulas with an earlier generic intersection silently lost later
`INDEX` attribution, and repeated computed intersections emitted only the first
event. A nested name such as `Both = LET(x,Pick B5,x)` then used in `Both:C6`
could persist two differently worded `I_DYNAMIC_REF` rows for the one INDEX.
The first memo implementation also let inherited asyncio tasks and copied
thread contexts mutate one dictionary; an owner check based on `get_ident()`
still failed when Windows recycled a terminated thread's id. Retained execution
objects make ownership unambiguous, while the context identity wrapper prevents
stale id reuse and preserves same-owner nested cache hits. Permanent tests cover
mixed and repeated intersections, nested diagnostic deduplication, child-task
and copied-thread isolation, and actual owner-object identity. Independent
stress passes exercised 33 intersections per formula, 1,024 inherited
executions, exception resets, recursive/depth-bounded aliases, and an exact
numeric-thread-id recycling reproducer without finding another defect. No
formal review invocation was spent in this hardening cycle.

## 2026-07-27 P3 — Initial formal split and declaration investigation

Decision: charge global formal invocation #17 / R-mech #10 as `REVISE` and
global invocation #18 / R-test #5 as a clean `APPROVE` on the initial frozen
tree, then reopen P3 because the mechanics finding required a runtime oracle.
Bind that historical candidate to base
`3359a974dc72db1dd1ec47507eaf24891c670c92`, stage tree
`4b8364b6060ec52245260448601c9c61f19856a3`, and cached-diff blob
`1a0b6eb4c92cc56345203f1a7f4171aab7c66628` (45 files, 14,600 insertions,
156 deletions). Do not carry the old test approval across any correction.

Alternatives considered: accept general Name Manager rules as the LET/LAMBDA
grammar without testing; treat the clean R-test verdict as approval of later
code; suppress R1C1-like declarations; permit periods because one Microsoft
page says names may contain them; or proceed to P4 with a split gate.

Rationale: Microsoft's LET, LAMBDA, and general formula-name pages disagree at
their boundaries. A 49-probe worksheet-entry matrix on desktop Excel 16.0 build
19530 produced 26 calculated formulas and 23 rejected literals: the UI accepted
Unicode, underscore/backslash, R1C1-like, and beyond-grid A1-like locals, while
rejecting in-grid A1 spellings, periods, operators, spaces, `@`/`#`, leading
digits, and same-scope case-insensitive duplicates. COM `Range.Formula2`
followed a distinct acceptance path. Saved OOXML then exposed the real gap:
Excel prefixes ordinary LET/LAMBDA declarations and their uses with `_xlpm.`.
The exact UI matrix is committed as
`docs/evidence/p3-excel-declaration-oracle.csv`; the unsaved synthetic Book1 is
left open pending explicit user authorization to discard it.

## 2026-07-27 P3 — Exact stored namespace and composite remediation

Decision: model raw and `_xlpm.` locals as distinct exact lexical namespaces;
strip `_xlpm.` only for public display and duplicate-name identity. Preserve
unprefixed built-in precedence in function-call position. Retain concrete
precedents plus conservative opacity for direct, reversed, and grouped lexical
colon/intersection endpoints; keep scalar locals non-reference-valued and
propagate reference-valued outer attribution. Reuse one matching-group map per
formula and skip scope-bound occurrences in the global intersection prepass.

Alternatives considered: strip `_xlpm.` for every lookup; expose it in public
function labels; let a prefixed local shadow an unprefixed built-in or defined
name; treat all lexical composites as exact; discard the operator; rerun group
matching for every whitespace token; or preserve the first formal candidate
because its full test suite was green.

Rationale: prefix-stripped lookup made `_xlpm.SUM` shadow raw `SUM` and leaked
private labels. The scope-blind intersection pass could resurrect a global
`Pick = INDEX(...)` beneath `LET(Pick,...)`. Parenthesized locals beside `:`
bypassed direct composite handling, while 1,900 grouped intersections rebuilt
the parenthesis map 1,900 times. Exact-style keys, intrinsic/display function
separation, occurrence-index skipping, and grouped endpoint recovery close
those gaps. The maximum-sized 7,615-character stress probe now completes in
0.099 seconds on the development machine. F19/F20 fixtures, semantic/map
goldens, oracle expectations, and map budgets were refreshed together.

The user explicitly authorized unbounded review loops and continued subagent
orchestration at high reasoning. This supersedes the earlier 30-invocation hard
stop without erasing its nominal allocation: every extra verdict remains
charged, every reviewer remains fresh/stateless, and every changed candidate
must receive new R-mech and R-test reviews before the P3 gate can close.

## 2026-07-27 P3 — Raw axis-shaped lexical precedence

Decision: when a simple raw spelling in a colon composite exactly matches an
active LET/LAMBDA binding, resolve that lexical binding before contextual
whole-column or whole-row recovery. Preserve known defined names and ordinary
unbound axis ranges, keep scalar bindings non-reference-valued, retain concrete
and dynamic identity for reference-valued bindings, and exclude qualified,
structured, external, and 3-D spellings from lexical reinterpretation.

Alternatives considered: classify every all-letter pair as an A1 whole-column
range before scope analysis; always prefer lexical-looking text even when no
binding exists; treat every ambiguous composite as reference-valued; or reject
the spelling instead of containing it conservatively.

Rationale: a read-only preflight proved that `r:s`, `rr:ss`, and grouped or
reversed variants could falsely promote scalar LET expressions and emit
`I_DYNAMIC_REF`. A disposable live-Excel differential confirmed lexical
precedence for reference-valued `C:C` and `R:R` locals and errors for scalar
locals. The corrected candidate passed 319 focused reference tests, a 576-case
LET/LAMBDA cross-product, a second independent matrix exceeding 1,000 cases,
and bounded 900- and 1,000-expression stress probes. The full repository then
passed 764 tests at 90.35% branch coverage. These were non-verdict preflights;
fresh formal R-mech and R-test reviews remain required after fingerprinting.

## 2026-07-27 P3 — Second formal split and contextual tile provenance

Decision: charge global invocation #19 / R-mech #11 as `REVISE` and global
invocation #20 / R-test #6 as a clean `APPROVE` on the second frozen tree, then
invalidate the test approval because the mechanics corrections changed that
tree. Bind the historical candidate to base
`3359a974dc72db1dd1ec47507eaf24891c670c92`, stage tree
`c7de5e88dbccc9b0c0f7276af1ed3bd585d8ee0a`, and cached-diff blob
`f4e58e1213c4b4616a3a8d9e05d6c05255df9eeb` (51 files, 15,923 insertions,
193 deletions).

Alternatives considered: inspect only complete structured tokens; reuse the
block-anchor formula at every tile; translate coordinate spills with ordinary
relative references; retain the initial anchor's static endpoints when grouped
operator folding changes by context; or omit opacity markers that appear only
in later tiles.

Rationale: R-mech #11 reproduced incorrect multi-table coverage for
`[@Input]:F2` and relative-endpoint drift after horizontal tiling. Structured
requirements now aggregate recursively across safe composite operands and can
retain multiple qualified current-row tables. Each tile analyzes the formula
translated from the block origin with coordinate spills deliberately fixed.
An independent 101-case block-versus-per-cell differential then found grouped
colon ownership and intersection-opacity mismatches at inside/outside table
boundaries. Treating opaque structured failures as foldable endpoints and
unioning tile-derived `opaque:ref` markers closed both gaps. The differential
finished with zero failures, the focused P3 slice passed 501 tests, and the full
repository passed 771 tests at 90.41% branch coverage. Fresh formal R-mech and
R-test verdicts remain required on the next exact fingerprint.

## 2026-07-27 P3 — Final clean approvals and gate close

Decision: close P3 after global invocation #21 / R-mech #12 and global
invocation #22 / R-test #7 both returned clean `APPROVE` verdicts with no
findings on one unchanged candidate. Bind the approved tree to base
`3359a974dc72db1dd1ec47507eaf24891c670c92`, stage tree
`a6ced2a07ca3e438c47e11e93e261b572699cc50`, and cached-diff blob
`fe3d0c40a75446da3d1b6af146292eae778eb3f6` (51 files, 16,309 insertions,
194 deletions; 107 tracked files; zero tracked forbidden junk).

Alternatives considered: carry forward R-test #6 from the prior tree; combine
the two domains; accept only the author's differential; or leave current-state
documentation gate-pending after formal closure.

Rationale: both fresh reviewers independently reproduced the 501-test focused
slice, all 771 repository tests, and 90.41% branch coverage. R-mech #12 also ran
a separate 605-case block-versus-per-cell differential over contextual
structured composites with zero failures. Each reviewer confirmed the stage
tree and cached-diff blob before and after inspection and found no unstaged or
untracked changes. The review ledger now records 12 R-mech, 7 R-test, and 3
R-repo invocations; P4 may begin only after the P3 milestone commit.

## 2026-07-27 P4 — Exact circular verification over block condensation

Decision: retain formula blocks as the persisted graph granularity and run an
iterative Tarjan pass over that condensed graph. Reanalyze only candidate SCCs
at concrete-cell anchors. Treat a singleton candidate as proven acyclic only
when every internal dependency has one strict row-major direction; otherwise
apply the same bounded exact fallback used for multi-block SCCs. Index coarse
block intersections and exact source ownership with balanced per-sheet
rectangle trees.

Alternatives considered: report every coarse self-overlap as circular; stop
after a singleton self-inclusion scan; allocate a complete 50,000-cell graph;
resolve every exact dependency by scanning all blocks; or trust anchor-only
geometry for structured, spill, and composite references.

Rationale: the F09b running total has one coarse self-overlapping block but is a
strict DAG, while a homogeneous singleton block can still contain an indirect
two-cell cycle. The monotonic proof separates these without quadratic work;
ambiguous cases retain the handoff's 64-seed/100,000-cell bounds and emit
`W_POSSIBLE_CIRCULAR` when proof is incomplete. Exact reanalysis preserves
current-row structured context, coordinate spill anchors, and composite hulls.
F09a emits one canonical `E_CIRCULAR`; Stage 2a visits all 50,000 F09b formula
cells exactly once and proves the strictly earlier dependency direction without
entering Stage 2b.

## 2026-07-27 P4 — Ranked bidirectional spatial traversal

Decision: assign dense ranks to unique public `GraphHop` ordering keys and
persist those ranks in both destination and source spatial mirrors. Use an
`rtree_i32` rank dimension with range-existence binary probes. For the interval
fallback, cache a deterministic revision-bound rectangle BVH and use
minimum-rank branch-and-bound search. Rebuild both mirrors atomically after any
formula refresh and reject dirty, partial, mixed, missing, orphaned, or
mismatched state as `E_CORRUPT`.

Alternatives considered: scan edge IDs until a cap is filled; sort all matching
fan-out in memory; assume internal edge-ID order matches public casefold order;
index destinations only; permit equal ranks without proving equal public hops;
or use the fallback B-tree's first matching rank scan.

Rationale: initial preflights reproduced linear work over irrelevant edges,
wrong capped prefixes after reversed edge IDs, a source-mirror blind spot, and
duplicate ranks that represented distinct public hops. Required dependent and
precedent semantic keys now enforce a rank/key bijection while still allowing
duplicate physical edges for the same public hop. RTree progress callbacks at
1,000/10,000/50,000 irrelevant edges remained 3/5/8 for dependents and 4/8/11
for precedents. The interval BVH's independent final-rank adversary visited
39/63/75 nodes. Both directions preserve their exact semantic prefix across
1,200 reversed physical IDs.

## 2026-07-27 P4 — Preflight hardening before the formal gate

Decision: keep all review findings as preflight until one staged fingerprint is
frozen. Export relational ranks, graph state, and both mirrors through natural
keys so full-versus-incremental canonical equality cannot pass vacuously. Add a
deterministic Hypothesis RTree/interval differential, a pinned 4,000-token graph
golden budget, an authorizer covering all five graph surfaces, and a persisted
clean-corruption matrix over both directions and backends.

Alternatives considered: count the first green 112-test slice as sufficient;
exercise mirror corruption only through the dirty gate; treat a byte-identical
golden as its own token-budget proof; retain a fixed three-edge spatial parity
example instead of §8.2 randomized coverage; or charge changing preflight trees
as formal verdicts.

Rationale: independent mutation probes showed the lifecycle export could omit
all new ranked state, the F09b test could pass with circular detection disabled,
and dirty-gate tests did not execute deeper mirror validation. The remediated
tests require exactly 50,000 Stage 2a resolutions, reject Stage 2b allocation,
compare 80 deterministic random edge sets with a brute oracle, and independently
kill source-only, rank-blind, missing-row, and relational-cross-check mutants.
The current focused P4 slice passes 141 tests. Fresh mechanics preflight returned
`APPROVE`; the test preflight finding is remediated and awaiting its final
closure message. No formal review invocation has yet been charged for P4.

## 2026-07-27 P4 — Preflight closure and freeze readiness

Decision: accept the remediated preflight candidate for fingerprinting only
after the test reviewer reran the exact surviving mutations and returned clean
`APPROVE`. Keep the formal phase gate open until two new stateless reviewers
inspect one unchanged staged fingerprint.

Rationale: the 16 clean-state mirror cases killed the complete-validator,
source-only, rank-blind, and relational-cross-check mutants exactly as intended.
The post-remediation focused slice passed 141 tests; the full repository passed
877 tests in 367.54 seconds at 90.15% branch coverage. Ruff, formatting,
Pyright, fixture generation, the parser oracle, lock validation, and package
build were also clean. These remain implementation evidence, not formal review
credits.

## 2026-07-27 P4 — First formal split exposes completeness and fixture-oracle gaps

Decision: charge global invocation #23 / R-mech #13 and global invocation #24 /
R-test #8 as `REVISE` on the first frozen P4 candidate. Bind that rejected
candidate to base `6eb092b12c1c0398ac76a6153c52096f08904d7a`, stage tree
`cab97b4fcd15f23b772b9ca2cc56978d44defd43`, and cached-diff blob
`cce5c2a7c1a449e4a6bafad2a81dc8fb1de6759e` (30 files, 7,544 insertions,
121 deletions). Reopen both formal domains after any correction.

Alternatives considered: carry the clean preflight verdicts into the formal
gate; treat trigger-maintained `dirty` as sufficient after an external writer
manually restores it to clean; perform a full relational edge scan before every
bounded query; accept F03's one golden path plus I12 round trips as proof that
no required edge is missing; or charge only the mechanics rejection.

Rationale: R-mech #13 deleted the queried-direction mirror row, restored only
the persisted clean bit, reopened the spatial facade, and obtained successful
empty precedent/dependent traces on both backends. It also removed
`graph_spatial_state` from a schema-v5 sidecar and reproduced a raw
`OperationalError` instead of rebuild or structured corruption. Any fix must
retain the demonstrated irrelevant-edge bound, so query-time relational scans
are not acceptable. R-test #8 independently deleted the required
`Summary!C7 -> Calc!B2` dependency; both the staged F03 golden and I12 loop
survived because neither enumerated an independently authored complete F03
edge set. Both reviewers rechecked the exact fingerprint after their probes and
left the worktree/index unchanged.

## 2026-07-27 P4 — Formal #13/#8 remediation

Decision: add persistent `mutation_epoch` and `clean_epoch` trust state rather
than scanning relational edges before bounded queries. Every edge, block,
sheet, destination-mirror, and source-mirror trigger increments the mutation
epoch and marks the graph dirty; a successful atomic rebuild alone seals the
current epoch. Validate the exact state-table columns and all 15 required dirty
triggers before accepting a current-version sidecar. Separately freeze F03 with
an independently authored 13-edge semantic projection keyed by source and
destination geometry plus `via`.

Alternatives considered: query-time `COUNT(*)` or relational overlap scans;
process-local trust only; trust the caller-restored dirty bit; add a second full
spatial mirror solely for verification; catch the missing table only at query
time; or expand the F03 trace golden without enumerating every semantic edge.

Rationale: epoch equality is an O(1) prerequisite and survives reopening, while
all expensive validation remains on mutation/rebuild boundaries. The deeper
per-returned-edge validators remain independently tested by forging both trust
fields. Missing or malformed state columns/triggers now cause a monotonic
disposable-sidecar rebuild; live database-shape/value errors are converted to
structured `E_CORRUPT`. The F03 oracle contains the eight Calc edges and five
Summary edges implied directly by generator-authored formulas, explicitly
including `Summary!C7 -> Calc!B2`; `Counter` equality rejects duplicates as well
as missing or extra edges. The remediated focused P4 slice passes 155 tests.

## 2026-07-27 P4 — Closure-preflight trust-state hardening

Decision: require complete canonical SQL equality for every persistent dirty
trigger, reject additional persistent triggers that can affect protected graph
tables or trust state, and validate relational rank aggregates once whenever a
current sidecar is opened. Both rank directions must be non-null, dense from
one, and match their persisted maxima. Shape direct `DependencyGraph`
construction failures as `E_CORRUPT`.

Alternatives considered: retain substring-based trigger inspection; accept a
stored maximum without comparing relational ranks; scan the complete graph
before every bounded trace; reject all connection-local TEMP triggers; or
recompute every semantic ordering key to defend against a caller that
coherently rewrites the relational catalog, both mirrors, and trust markers.

Rationale: an independent mechanics preflight used comments, `WHEN 0`, undo
statements, and extra persistent triggers to keep stale mirrors falsely clean;
it also lowered each stored maximum so bounded traces silently suppressed valid
edges. Exact trigger definitions and open-time aggregate validation close those
persisted-sidecar failures without changing trace work. Connection-local TEMP
triggers and a fully self-consistent malicious rerank remain outside the frozen
sidecar-corruption contract because they require a caller already controlling
the live internal connection or an integrity root outside the same database.
The closure reviewer explicitly reproduced and accepted that boundary, then
returned `APPROVE`. Fresh author verification passed the 173-test focused P4
slice in 84.49 seconds. This approval is preflight evidence only; the changed
candidate still requires new stateless formal R-mech and R-test verdicts.

## 2026-07-27 P4 — Second formal split finds bounded-integrity and oracle gaps

Decision: charge global invocation #25 / R-mech #14 and global invocation #26 /
R-test #9 as `REVISE`. Bind the rejected candidate to base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, stage tree
`dd616019ff38c4946aed8773b106225f2f16c673`, and cached-diff blob
`8dd7ff01c0188e5f210b04d1adf7dce497ee1193` (31 files, 8,346 insertions,
128 deletions). Reopen both formal domains after remediation.

Alternatives considered: treat the prior preflight approval as closing the
stronger forged-epoch case; rely on trigger state without detecting an absent
queried mirror; allow raw `OperationalError` for malformed interval columns;
call set equality an exact edge oracle; or charge only the mechanics verdict.

Rationale: R-mech #14 deleted the destination mirror for a real whole-column
edge, forged both the clean bit and epoch, and obtained a successful empty
bounded dependent trace on both backends. The same retrieval structure affects
precedent traces and paths. It also removed the interval `rank` column while
preserving canonical triggers; current-schema validation accepted the sidecar
and opening leaked `no such column: rank`. R-test #9 independently duplicated
every F05 structured-reference edge and added a valid unexpected singleton-cell
edge; all 20 real-fixture tests survived because the F04/F05/F15/F19 projection
used a set and an inner fblock join. Both reviewers confirmed the unchanged
fingerprint at exit. The rejected tree nevertheless passed 909 tests and
90.09% branch coverage; green evidence does not override surviving mutations.

## 2026-07-27 P4 — Formal #14/#9 remediation and adversarial closure

Decision: validate both complete mirrors and exact physical storage whenever a
persisted sidecar opens, then seal the accepted mutation epoch in process for
O(1) live checks. Share one canonical rank projector between rebuild and open
validation so persisted ranks must match the exact public semantic order. When
corrupt RTree shadow storage prevents valid DDL teardown, build and checkpoint a
same-directory replacement and install it atomically. Require graph rebuilds to
run in store-owned transactions. Replace every graph-fixture set oracle with an
independently authored lossless multiset.

Alternatives considered: scan relational edges before every bounded query; add
more caller-forgeable persisted counters; validate only mirror row counts;
duplicate canonical rank logic in the validator; use `writable_schema` to excise
corrupt RTree objects; delete and recreate the database with a no-file window;
attempt to adopt raw SQLite transactions without commit/rollback callbacks; or
retain set-based fixture projections after fixing only F03.

Rationale: the first closure preflight proved that a forged clean epoch could
hide missing or displaced queried mirrors. Full open-time bidirectional mirror
comparison closes persisted damage, while a process seal detects post-open
mutation without changing trace work. Subsequent adversarial passes changed
only `via`, coherently swapped relational and dual-mirror ranks, deleted RTree
node/rowid storage, forced atomic-replacement failures, and rolled back a raw
outer transaction. The shared projector now rejects noncanonical semantic ranks;
RTree integrity probes enter a generation-preserving atomic recreation path;
replacement failures preserve the original file and error; and raw transaction
adoption is refused before context entry or record consumption. Managed commit,
rollback, nesting, seal publication, and interval-cache restoration remain
green. Independently authored F03/F04/F05/F15/F19 Counters reject missing,
duplicate, opaque, malformed, and unexpected singleton-cell edges. Final
non-formal mechanics and test closure checks returned `APPROVE`; they do not
count as formal invocations. Fresh author verification passes the 214-test P4
slice in 71.80 seconds. A new unchanged fingerprint still requires fresh
stateless R-mech and R-test formal verdicts.

## 2026-07-27 P4 — Third formal split finds failed-commit and live-error gaps

Decision: charge global invocation #27 / R-mech #15 as `REVISE` and global
invocation #28 / R-test #10 as `APPROVE` on the third frozen candidate. Bind the
rejected candidate to base `6eb092b12c1c0398ac76a6153c52096f08904d7a`, stage
tree `c5d870a081a3c67a87b2462fd5ae763be2625a77`, and cached-diff blob
`261abc57cf92a208bca22c52cbb6e2c2e7ab6b78` (31 files, 9,427 insertions,
141 deletions). A changed mechanics fix requires fresh reviewers in both
domains despite the clean R-test verdict.

Alternatives considered: carry R-test #10 onto a changed tree; treat SQLite
commit failure as caller recovery; rely on process-seal equality to make every
live RTree read safe; catch only the exact reviewer query; or charge only the
rejecting verdict.

Rationale: R-mech #15 enabled deferred foreign keys so `commit()` failed after
the managed body. The connection remained inside the failed transaction, the
invalid row remained visible, and the edge-store seal still marked it as an
active trusted outer transaction; a later store context entered as nested.
Separately, deleting active RTree node storage and restoring the exact sealed
epoch caused `direct_dependents` to leak raw `sqlite3.DatabaseError` from its
relational/spatial validation join. Managed commit failure must rollback and
restore seal/cache before re-raising, and every public graph surface must shape
live SQLite storage failures as `E_CORRUPT`. R-test #10 independently killed
duplicate, singleton, opaque, and orphan fixture mutations and passed all 958
tests, but the exact staged tree is still rejected. Both reviewers confirmed the
unchanged fingerprint at exit.

## 2026-07-27 P4 — Formal #15 remediation and transaction-finalizer closure

Decision: make graph transaction cleanup explicit and idempotent with no-I/O
commit and rollback finalizers, invoke the matching finalizer as a backstop even
when a lifecycle hook fails before or after its own state transition, and shape
genuine live SQLite storage failures on all five public graph-query surfaces as
`E_CORRUPT` while preserving `sqlite3.ProgrammingError` and unrelated exception
identity. Require fresh stateless mechanics and test reviews after the final
verification freeze.

Alternatives considered: restore the process seal only in the ordinary rollback
hook; attempt rollback after a successful SQLite commit when its publication
hook fails; let hook failures replace the primary body or commit failure; catch
the broad `sqlite3.DatabaseError` hierarchy without a caller-misuse carveout; or
reuse the clean R-test #10 verdict on the changed implementation.

Rationale: a deferred-constraint commit failure must rollback the connection and
restore the old seal before control returns, whereas a post-commit hook failure
must publish the new seal and clear bookkeeping without undoing durable data.
Hooks can themselves fail before or after their state transition, so the store
now applies a corresponding no-SQL finalizer unconditionally and preserves the
primary error while chaining cleanup failures. A focused matrix covers managed
body, commit, rollback-hook, and commit-hook failures on both sides of the state
transition; repeated finalization performs no database I/O. Query regressions
cover corrupt live RTree storage plus `ProgrammingError`, `TypeError`,
`ValueError`, and existing `ExcelLSPError` identity on every public surface.
Independent non-formal mechanics and bounded-query closure reviews both returned
clean `APPROVE`; those closure checks are not verdict-bearing invocations. Fresh
author verification passes the 273-test focused P4 slice in 73.83 seconds. The
candidate remains gate-pending until a new exact staged fingerprint receives
fresh formal R-mech and R-test verdicts.

## 2026-07-27 P4 — Fourth formal split finds live-maxima and exact-traversal gaps

Decision: charge global invocation #29 / R-mech #16 and global invocation #30 /
R-test #11 as `REVISE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`9a242fb56398bbf11ecf5f6d15861f2ec857c9dd`, and cached-diff blob
`cfeeefcfca91fad93cda73cba9384f5c96907e11` (31 files, 10,250
insertions, 146 deletions). Reopen both formal domains after root-cause
remediation under the unbounded-review authorization.

Alternatives considered: treat mutable rank maxima as harmless metadata; rely
on direct-query edge validation to protect bounded traces; refresh the facade
after every live sidecar mutation; accept a 50,000-call count plus first/last
cells as an exactly-once traversal proof; or carry either rejected verdict onto
a changed tree.

Rationale: R-mech #16 lowered `dependent_rank_max` and
`precedent_rank_max` after opening, without changing the three fields held in
the process seal. Cached dependent and precedent traces on both spatial
backends then silently returned one node with `truncated=false` even though the
corresponding direct query still found a real edge. Every persisted graph-state
field used by a query must therefore participate in the live trusted tuple and
transaction snapshot. R-test #11 replaced one interior F09b cell with a
duplicate neighbor while preserving 50,000 resolver calls plus the expected
first and last cells; all 19 circular/F09b tests still passed, leaving a real
self-reference at the omitted cell potentially undetected. The boundedness
guard must assert the complete ordered 50,000-cell sequence. Both reviewers ran
the 273-test P4 slice and all 1,009 repository tests successfully, confirmed the
unchanged fingerprint at exit, and made no workspace edits; green baseline
tests do not override either surviving adversarial mutation.

## 2026-07-27 P4 — Formal #16/#11 remediation and complete live-state closure

Decision: replace the mutation-epoch-only seal with one immutable seven-field
`graph_spatial_state` tuple, return that validated tuple from `require_clean()`,
and consume its direction-specific maximum without a second database read.
Strengthen F09b's boundedness guard to compare the complete ordered list of
50,000 expected cells. Keep the candidate gate-pending for fresh formal reviews.

Alternatives considered: seal only the two maxima in addition to the epoch;
revalidate full mirrors on every bounded hop; retain a second maximum query
after validating the tuple; cover only the reviewer-supplied lowered-maximum
cases; or assert F09b set equality without order and multiplicity.

Rationale: the complete immutable state holds singleton, dirty, dependent and
precedent rank maxima, revision, mutation epoch, and clean epoch. Initial open,
transaction snapshot, pending rebuild, durable-commit publication, rollback
restoration, and idempotent no-I/O finalizers all carry that same value. Cached
facade regressions now alter every mutable state field or delete the row on both
RTree and interval backends; each public query returns structured `E_CORRUPT`.
The direction-specific maximum cases additionally reproduce both precedent and
dependent traces. F09b now requires exact equality with
`CellNode(1, row, 2)` for every row from 3 through 50,002 while retaining its
Stage-2b hard-failure sentinel and monotonic proof. Focused remediation checks
passed 18 cached-state cases, both real-F09b paths, the 12-test governance
contract, Ruff, formatting, Pyright, and whitespace validation. These are
author/remediation checks, not formal verdicts; a new frozen fingerprint still
requires fresh stateless R-mech and R-test reviews.

## 2026-07-27 P4 — Fifth formal split finds an operation-snapshot race

Decision: charge global invocation #31 / R-mech #17 as `REVISE` and global
invocation #32 / R-test #12 as a clean `APPROVE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`904516baeb260ed2dc7a552037230a4f907f0c37`, and cached-diff blob
`c9ac51420d96b65cd5c4520edefab08fec7e8f07` (31 files, 10,389
insertions, 146 deletions). Reopen both domains after mechanics remediation.

Alternatives considered: treat concurrent sidecar mutation as unsupported;
recheck only the two direct methods on exit; accept an empty result when the
dirty marker changes during retrieval; carry the clean R-test #12 verdict onto
a changed tree; or rely on the seven-field entry seal without a stable SQLite
snapshot.

Rationale: R-mech #17 used a second connection to delete the queried mirror row
after entry validation but before spatial retrieval, and to atomically delete
the source mirror plus relational edge before precedent retrieval. Direct
dependent and precedent queries on both RTree and interval backends silently
changed from one result to zero while `dirty=1` and mutation/clean epochs
diverged. A trace or path can likewise mix generations. Every complete public
graph operation therefore needs one consistent read snapshot with trust
validation inside it, without committing or rolling back a caller-owned
transaction. R-test #12 independently matched 5,000 randomized circular graphs
to a brute oracle, matched 1,000 ranked spatial queries across both backends,
and passed the 291-test P4 slice, 21 fixture/oracle tests, and all 1,027 tests at
90.15% coverage with no findings. Both reviewers confirmed the exact unchanged
fingerprint and no workspace edits; the test approval remains charged but
cannot approve the forthcoming changed tree.

## 2026-07-27 P4 — Formal #17 remediation with operation-wide read snapshots

Decision: wrap all five public graph operations in one deferred SQLite read
snapshot whenever no caller transaction already exists, place the complete
trust-state check as the snapshot's first read, and release only graph-owned
snapshots on every success or exception. Preserve raw and store-managed caller
transaction ownership and all existing exception identity rules.

Alternatives considered: revalidate only on method exit; take one snapshot per
hop; use `BEGIN IMMEDIATE` and block writers before validation; always rollback
the connection even when the caller owns its transaction; or handle only the
two direct surfaces from the reviewer reproduction.

Rationale: a deferred read transaction provides one coherent database version
without acquiring a write lock. In WAL mode, a second connection can commit a
new coherent graph after the trust read while the current operation continues
to see the complete prior version. In DELETE rollback-journal mode, the writer
can prepare its change but its commit waits until the reader releases, after
which a fresh facade sees the new version. Every public direct, trace, and path
surface now enters that boundary. Graph-owned snapshots roll back on success,
`DatabaseError`, `ProgrammingError`, invalid input, and unrelated exceptions;
cleanup keeps the primary exception and chains cleanup failures. Existing raw
and managed transactions remain active and retain caller work. Initial focused
remediation passed 40 race/ownership/error cases and the broad graph/index slice
passed 259 tests in 52.58 seconds. Independent external probes passed all 20
surface/backend/journal race combinations and 30 transaction/error ownership
combinations. Their durable checked-in forms pass 60 snapshot/ownership tests,
including event-synchronized rollback-journal commit blocking and every caller
transaction/error combination; the complete focused P4 slice passes 351 tests
in 76.78 seconds. No formal verdict is implied by this remediation entry.

## 2026-07-27 P4 — Sixth formal split finds snapshot-release lock leakage

Decision: charge global invocation #33 / R-mech #18 as `REVISE` and global
invocation #34 / R-test #13 as a clean `APPROVE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`15bd6efc3d6bc7ba7dd06d16627f98d288ce009a`, and cached-diff blob
`9448472cc182a4b7bc907063bb7c3122035bb3e9` (31 files, 10,789
insertions, 146 deletions). Reopen both domains after cleanup remediation.

Alternatives considered: leave failed rollback recovery to `close()`; regard a
rollback exception as proof the transaction ended; preserve the primary error
without restoring connection state; make future queries inherit the leaked
transaction as caller-owned; or carry the clean R-test #13 verdict forward.

Rationale: R-mech #18 denied `SQLITE_TRANSACTION/ROLLBACK` in DELETE journal
mode. A successful query's graph-owned rollback raised, left
`connection.in_transaction=true`, and caused a second writer's commit to fail
with `database is locked`; a later query mistook the leaked transaction for a
caller transaction and kept the obsolete snapshot. An injected primary
`ValueError` preserved its identity and chained the cleanup error but leaked the
same lock. Snapshot release therefore needs direct-SQL fallback and must close
or poison the connection if release remains impossible. `IndexStore.close()`
must likewise close in a `finally` path and finalize graph bookkeeping despite
rollback failure. R-test #13 killed snapshot, revision-seal, skipped-F09b-cell,
and raw-transaction-refusal mutants; its 351-test P4 review found nothing else.
Both reviewers ran the 1,087-test repository and confirmed the unchanged frozen
scope; the test approval is charged but cannot approve a changed cleanup tree.

## 2026-07-27 P4 — Formal #18 remediation and conclusive snapshot release

Decision: centralize graph-owned snapshot release with rollback-state
inspection, direct-SQL `ROLLBACK` fallback, and connection poisoning plus
verified physical close when release remains uncertain. Harden `IndexStore`
close and context exit with the same layered recovery, unconditional no-I/O
finalization, verified retryable close state, and exact primary-error precedence.

Alternatives considered: retry rollback once; mark a graph poisoned while
leaving its physical connection open; set `IndexStore._closed` before verifying
the handle; discard earlier body causes when cleanup also fails; assume a close
override always takes effect before raising; or omit rollback-journal writer
probes after WAL passed.

Rationale: recoverable rollback failures now fall through to direct SQL and
leave the connection reusable. If rollback and fallback both fail, every graph
surface poisons the facade and closes the native SQLite descriptor; after three
failed overridden/direct close attempts, the native base descriptor provides a
supported emergency release. `IndexStore.close()` independently retries and
verifies physical closure, stays retryable if closure cannot be proven, always
runs the graph rollback finalizer, and always retains the context body as the
primary exception. Existing explicit causes or implicit contexts are grouped
with cleanup failures after removing back-links to the primary, so traceback
chains remain complete and acyclic. Checked-in matrices cover every public
surface, both backends, WAL/DELETE, successful and failing bodies, fallback and
denied rollback, raw/managed ownership, transient close-before/after behavior,
native subclass overrides, writer unblocking, store reopen/trust, hooks, and
finalizers. Independent adversarial closure passed 80 native close cases,
96 context-close cases, 32 reopen/trust cases, and 36 hook/finalizer cases with
no remaining blocker or minor. Fresh author verification passes the 557-test P4
slice in 79.31 seconds. These remediation results are non-verdict evidence; a
new exact fingerprint still requires fresh stateless reviews in both domains.

## 2026-07-27 P4 — Seventh formal split finds acquisition and native-close gaps

Decision: charge global invocation #35 / R-mech #19 and global invocation #36 /
R-test #14 as `REVISE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`166634a31d389df75b47122aea863571640efb35`, and cached-diff blob
`24a3c53ae333be3ba684d9c8f32f6585d544d6b6` (31 files, 11,785
insertions, 152 deletions). Reopen both domains after root-cause remediation.

Alternatives considered: regard a raised `BEGIN` as proof no transaction began;
leave native-subclass emergency closure exclusive to graph queries; accept
virtual close retries without proving physical release; replace a transaction
body's prior cause with cleanup evidence; carry any earlier test approval onto
the changed tree; or defer the missing matrices to a later phase.

Rationale: R-mech #19 used real SQLite connection subclasses whose `execute`
performed `BEGIN` or `BEGIN IMMEDIATE` and then raised. Both graph snapshot and
managed-store acquisition occurred outside their cleanup boundaries, leaving
an active transaction, inconsistent EdgeStore bookkeeping, and a blocked
rollback-journal writer. The same review showed that managed cleanup replaced
an existing explicit cause and left a cleanup-to-primary back-link, producing
a causal cycle. Both reviewers independently reproduced the adjacent
`IndexStore.close()` defect: three virtual close failures never reached the
native base descriptor, so denied rollback left the handle usable and a DELETE
writer locked. R-test #14 also established that the existing checked-in native
subclass coverage exercised `DependencyGraph`, not `IndexStore.close()` or
`__exit__`. Fresh reviewer verification otherwise passed the 557-test P4 slice
and all 1,293 repository tests, with clean Ruff, formatting, Pyright, lock, and
cached-diff checks. Both reviews confirmed the unchanged frozen fingerprint and
made no workspace edits.

## 2026-07-27 P4 — Formal #19/#14 remediation and acquisition-boundary closure

Decision: move graph and managed-store transaction acquisition inside their
respective cleanup ownership, give `IndexStore` the supported native SQLite
base-descriptor emergency close already used by graph snapshots, and preserve
the complete acyclic causal record for managed transactions and context exit.
Add real-connection fault matrices before freezing another candidate.

Alternatives considered: handle only the reproduced DELETE-journal cases;
assume `execute("BEGIN")` and `close()` are all-or-nothing; test hooks instead
of real `sqlite3.Connection` subclasses; close the connection without proving
descriptor state; retain primary identity while discarding its earlier cause;
or accept broad existing coverage without the exact reviewer reproductions.

Rationale: acquisition now enters the protected scope before `BEGIN` or `BEGIN
IMMEDIATE`, so both before-effect and after-effect failures run rollback-state
inspection, direct-SQL fallback, bookkeeping finalization, and conclusive close
when release cannot be proven. `IndexStore.close()` makes bounded virtual
attempts, checks physical state even when an attempt raises after taking effect,
then invokes `sqlite3.Connection.close(connection)` for a supported native
subclass whose override never reached the descriptor. Managed cleanup snapshots
the body's prior explicit cause or unsuppressed context, removes cleanup
back-links, groups distinct evidence once, and keeps the exact body exception
primary without cycles; context exit follows the same discipline.

The checked-in fault surface covers all five graph operations, both spatial
backends, WAL/DELETE, before/after-effect acquisition, normal/denied rollback,
managed acquisition, persistent close-before-effect, close-after-effect,
successful and failing context bodies, prior cause/context, writer unblocking,
and same-database reopen/trust. The focused fault matrix passes 142 tests in
4.98 seconds; the complete P4 slice passes 699 tests in 83.34 seconds. Fresh
repository verification passes all 1,435 tests in 311.74 seconds at 90.08% core
branch coverage. The exact F09b guard passes in 20.14 seconds, fixture/oracle
verification passes 21 tests in 18.44 seconds, and the standalone generator
reproduces all 14 current workbooks. These are remediation results, not formal
approval; both domains require new stateless verdicts on the next fingerprint.

## 2026-07-27 P4 — Eighth formal split finds virtual-state spoofing and evidence aliasing

Decision: charge global invocation #37 / R-mech #20 as `REVISE` and global
invocation #38 / R-test #15 as a clean `APPROVE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`a269b73c0575b8aa12f83a8b333caac0719e52da`, and cached-diff blob
`7a227abdaac35a0460e8646c03d6ef07fe005036` (31 files, 12,362
insertions, 153 deletions). Reopen both formal domains after remediation.

Alternatives considered: trust a closed-looking `ProgrammingError` from a
subclass override as native descriptor proof; treat a physically closed handle
as open when the same virtual property raises a different message; accept the
same exception identity twice because both semantic roles are meaningful; or
carry R-test #15 onto the changed closure and causal-composition tree.

Rationale: R-mech #20 built supported real `sqlite3.Connection` subclasses that
overrode both `close()` and `in_transaction`. A false closed-style state error
made graph and store cleanup return before native closure, leaving a DELETE
reader transaction and blocked writer. A false non-closed message after real
closure made `IndexStore._closed` remain false and exposed a dead connection as
usable. Native closure proof must therefore bypass virtual dispatch. The same
review reused one prior body cause as a rollback or grouped cleanup error; graph
snapshot cleanup, managed cleanup, and context exit each inserted that identical
object again as earlier evidence, so identity traversal visited it twice. The
composition must retain all evidence roles once by recursive identity.

R-test #15 independently returned no blocker, major, or minor finding. It ran
the 699-test P4 slice, exact 142-case native fault matrix, exact F09b guard,
1,435-test coverage suite at 90.08%, fixture/oracle checks, and all static checks,
and confirmed deterministic fixtures and the unchanged frozen scope. Its clean
approval remains charged but cannot approve a changed mechanics candidate.

## 2026-07-27 P4 — Formal #20 remediation with non-virtual proof and unique evidence

Decision: verify supported native SQLite descriptor state through the base
`sqlite3.Connection.in_transaction` descriptor, retain the existing virtual
path only for non-native proxies, and compose prior body evidence with cleanup
evidence through recursive identity containment. Normalize causal links between
members of graph cleanup groups before another formal freeze.

Alternatives considered: treat any virtual closed-style `ProgrammingError` as
proof; execute probe SQL that an authorizer could reject; trust the last close
call's return or exception; deduplicate only direct group members; leave an
earlier body cause attached after also placing it in the cleanup cause; or test
only the single-cleanup-error case rather than the reviewer's grouped alias.

Rationale: the base descriptor bypasses subclass properties without performing
SQL. It therefore sees the native transaction state while the handle is live
and the native closed-database error after physical closure, regardless of what
the virtual property reports. Graph poison cleanup and `IndexStore.close()` use
that proof after every virtual attempt and after the emergency base close.
Recursive containment prevents re-inserting an earlier cause/context already
inside cleanup evidence; once represented there, the obsolete primary link is
cleared before the cleanup cause is installed. Context-exit composition occurs
outside its cleanup handler so Python cannot auto-attach the cleanup exception
again. Graph multi-error groups additionally remove sibling causal links before
composition.

Real native subclasses now lie about state in both directions while failing
close before or after effect. The 44-case focused matrix covers all five graph
surfaces, both spatial backends, direct close, successful and failing context
exit, native usability, physical closure, DELETE writer release, `_closed`, and
same-database reopen. Its grouped-alias cases reuse the identical earlier cause
or implicit context as one member of a multi-error cleanup group across graph,
managed transaction, and context exit, and reject any repeated identity while
traversing every cause, context, and group member. The matrix passes in 1.93
seconds, graph plus IndexStore pass 681 tests in 64.24 seconds, and the complete
P4 slice passes 743 tests in 83.95 seconds. Fresh full verification passes all
1,479 tests in 324.82 seconds at 90.15% core branch coverage; exact F09b passes
in 21.27 seconds, fixture/oracle passes 21 tests in 20.34 seconds, and all 14
current fixtures regenerate deterministically. These are remediation results,
not formal approval; both domains require fresh verdicts on the next exact tree.

## 2026-07-27 P4 — Ninth formal split finds recursive aliases and rank-key splits

Decision: charge global invocation #39 / R-mech #21 as `REVISE` with two
minors and global invocation #40 / R-test #16 as a clean `APPROVE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`69226e78e489a34757b7bff6611c79cfe470c942`, and cached-diff blob
`7d3fc2c6e18d46c3a8ae29572bf8c3989f0362c7` (31 files, 12,829
insertions, 153 deletions). Reopen both domains after remediation.

Alternatives considered: accept the mechanics findings as non-corrupting edge
cases under the handoff's minor release valve; normalize only direct sibling
links; trust one `LIMIT 1` representative for every dense rank; rely on a fresh
facade to reject tampering; or carry R-test #16 to a changed integrity tree.

Rationale: R-mech #21 nested the earlier body exception inside an inner cleanup
group and pointed a separate outer member's suppressed context to that same
object. Direct-sibling normalization missed the cross-level alias, and complete
all-link traversal visited the identity twice across graph snapshot, managed
transaction, and context exit for both explicit and implicit body chains.
Normalization must use recursive group-membership closure.

The same reviewer duplicated one valid edge so both copies shared a canonical
semantic rank, cached the graph and exact process seal, changed one copy's
`via` to a distinct public hop, then restored all seven sealed state fields.
Direct dependents saw both hops, while bounded dependent trace and path selected
only the first edge at that rank and silently missed the split; a fresh
`EdgeStore` correctly rejected the noncanonical ranks. Every selected rank must
therefore prove that all edges it represents share one canonical public-hop key
without bulk materialization or unrelated-rank scans.

R-test #16 independently found no blocker, major, or minor. It passed the
743-test P4 slice, 142-case acquisition/close matrix, 44-case spoof/group matrix,
exact F09b guard, 1,479-test coverage run at 90.15%, two hash-seed golden probes,
fixture/oracle checks, and all static checks, and confirmed the unchanged scope.
Its clean approval remains charged but cannot approve a changed candidate.

## 2026-07-27 P4 — Formal #21 remediation with recursive ownership and rank identities

Decision: normalize cleanup evidence against complete recursive group ownership
and add an atomically rebuilt, trigger-invalidated `graph_rank_keys` catalog
that proves the canonical identity of each selected dense rank in O(1).

Alternatives considered: flatten every exception group and lose its structure;
strip only unsuppressed contexts; scan every duplicate relational edge on each
bounded hop; retain one representative without a persisted identity; invalidate
only relational mutations; or accept a fresh-facade rejection while cached
bounded queries remained inconsistent.

Rationale: recursive normalization first claims every nested group member, then
walks all cause and context subtrees and removes any link whose identity is
already group-owned. This includes suppressed contexts, preserves member order
and distinct evidence, and applies at graph snapshot, managed transaction, and
context exit boundaries.

`graph_rank_keys` is a `WITHOUT ROWID` table keyed by `(direction, rank)` with
one canonical serialized hop key. Rebuild validates rank identity, completes
both spatial mirrors, then repopulates the catalog before sealing the graph.
Edge, fblock, sheet, catalog, and active RTree/interval source/destination
mirror triggers invalidate an affected rank or direction. Open-time validation
checks exact catalog DDL, canonical contents, all 18 trigger bodies, and full
mirror integrity. Bounded traces and paths retrieve one key through the
composite primary key and compare it with the representative hop, so restoring
the seven state fields cannot restore a deleted split-rank identity. Valid
duplicates still share one rank, one catalog row, and one emitted hop.

The 24-case exact matrix covers recursive explicit/implicit body chains across
all three cleanup boundaries, suppressed cross-level contexts, dependent and
precedent rank splits, all three bounded surfaces, both backends, valid
duplicates, primary-key query planning, and active mirror invalidation. It
passes in 1.78 seconds. Fresh acquisition/close and native-state matrices pass
142 tests in 5.20 seconds and 44 tests in 1.89 seconds. The full P4 slice passes
769 tests in 108.32 seconds; all 1,505 repository tests pass in 335.57 seconds
at 90.19% core branch coverage. Exact F09b passes in 19.86 seconds,
fixture/oracle passes 21 tests in 18.62 seconds, and all 14 current fixtures
regenerate deterministically. These are remediation results, not formal
approval; both domains require new stateless verdicts on the next fingerprint.

## 2026-07-27 P4 — Tenth formal split finds residual cleanup aliases and an incomplete live seal

Decision: charge global invocation #41 / R-mech #22 and global invocation #42 /
R-test #17 as `REVISE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`94194a52eea6e62cc5a675cbcbdf9e3d3dcee719`, and cached-diff blob
`fc521684b55fd6df908e5acd6a0a3109962a5638` (31 files, 13,606
insertions, 252 deletions). Reopen both domains after remediation.

Alternatives considered: apply recursive normalization only to the three
previously tested boundaries; accept fresh-open rejection as sufficient for a
live cached facade; treat coordinated ordinary SQL and DDL changes as outside
the exposed-connection threat model; rely on implementation inspection for the
rank catalog's exact physical identity; or carry either verdict to a changed
tree.

Rationale: R-mech #22 reproduced repeated exception identities across three
remaining compositions: successful graph-body snapshot release with multiple
cleanup failures, direct `IndexStore.close()` aggregation, and successful
SQLite commit followed by both hook and finalizer failures. In each case a
later nested group's suppressed context pointed to the earlier error, so a
complete cause/context/group traversal reached the same identity twice. Every
multi-cleanup boundary must share the closure-aware normalizer.

The same reviewer coherently changed a duplicate-ranked relational edge and
its catalog row, restored the exact seven-field graph seal, and obtained a
stale bounded result from a cached `EdgeStore` on both backends. Dropping an
invalidation trigger before changing the duplicate edge produced the same
split. Fresh construction correctly rejected both states, but the live facade
must also bind its process-local rank identities and schema/storage tokens so
ordinary SQL or DDL cannot silently change its trusted graph.

R-test #17 found a separate major proof gap: permanent tests checked catalog
presence, one update trigger, primary-key lookup, and missing-table rebuild,
but did not kill relaxed exact-DDL, stored-content, and all-trigger validators.
Remediation must cover malformed physical DDL, missing/wrong/extra catalog
rows with the exact state seal restored, every catalog and contributing-table
invalidation family, direct rejection, monotonic current-sidecar rebuild, and
exact post-rebuild content on both backends. The reviewers otherwise passed
the 769-test P4 slice, all 1,505 repository tests at 90.19% core branch
coverage, exact F09b, fixture/oracle, acquisition/close, recursive-group, and
static-check matrices without mutating the frozen fingerprint.

## 2026-07-28 P4 — Formal #22/#17 remediation with process-local live identity

Decision: retain the exact persisted rank catalog, add process-local monotonic
live-storage witnesses, route every remaining multi-cleanup composition through
recursive identity normalization, and add permanent exact-DDL/content/trigger
kill matrices before requesting another formal verdict.

Alternatives considered: persist another restorable catalog revision; scan all
duplicate edges on each bounded query; accept fresh-facade rejection while a
cached facade remains stale; treat coordinated same-connection or cross-
connection SQL as impossible; ignore other-handle commits that restore all
persisted fields; or update only the tests named directly by the reviewers.

Rationale: a persisted nonce is as restorable as the seven graph-state fields.
The live facade instead seals the same handle's monotonic SQLite
`total_changes`, other-handle `data_version`, `schema_version`, and the complete
canonical rank-key map in process memory. Selected-rank validation remains one
composite-primary-key lookup. Managed commit publishes current live baselines;
rollback restores semantic/catalog snapshots but rebases `total_changes`
because SQLite counts rolled-back DML monotonically. Exact relational, catalog,
and seven-field restoration is now rejected on the originating handle and from
a second connection, as is trigger removal before a silent split. Fresh opening
still performs the complete canonical relational, mirror, catalog, DDL, and
trigger validation.

The exception normalizer now claims only recursive group membership—not unique
external causal subtrees—then removes every cause or context link, including a
suppressed context, whose identity is already group-owned. Successful snapshot
multi-cleanup detaches the primary from the later cleanup graph; direct close
normalizes its aggregate; and commit-hook/finalizer composition detaches and
normalizes before chaining. Ten snapshot surfaces plus both-backend direct-close
and post-commit cases traverse every cause, context, and group member and require
unique identities.

The new 52-case rank-catalog file proves exact columns, checks, composite primary
key, and `WITHOUT ROWID`; rowid-backed, extra-column, and wrong-constraint DDL;
missing, wrong, and extra catalog rows after exact seven-field restoration; all
18 trigger definitions individually missing and malformed; monotonic current-
sidecar reconstruction; exact canonical post-rebuild catalog content; and
functional INSERT/UPDATE/DELETE invalidation across edge, fblock, sheet,
catalog, and both active mirror families on both backends.

Fresh author evidence passes 777 graph/IndexStore/catalog tests in 92.14
seconds, the complete 839-test P4 slice in 126.75 seconds, and all 1,575
repository tests in 373.25 seconds at 90.15% core branch coverage. The focused
cleanup/acquisition matrix passes 156 tests in 9.09 seconds, the live-seal and
catalog matrix passes 62 in 7.44 seconds, exact F09b passes in 24.36 seconds,
and fixture/oracle passes 21 in 22.06 seconds. All 14 fixtures regenerate
deterministically; Ruff, formatting, Pyright, lock, diff, sdist, and wheel
checks are clean. These are remediation results, not formal approval; both
domains require new stateless verdicts on the next exact fingerprint.

## 2026-07-28 P4 — Pre-freeze closure of commit, constructor, and causal-identity gaps

Decision: reopen the candidate after two independent non-formal preflights found
three mechanics defects and five test-evidence weaknesses. Replace the coarse
same-handle witness with an authorizer-backed graph-write epoch, validate and
seal under one constructor-owned immediate transaction, rebuild immutable
exception groups when a chained primary is a member, assign shared external
causal identities one owner, and strengthen the catalog oracle before the next
formal freeze. These preflights are not verdict-bearing and do not change the
42-invocation review ledger.

Alternatives considered: retain `total_changes` and conservatively invalidate
every metadata-only transaction; rebase the monotonic witness after commit;
trust a deferred read snapshot to pair validated rows with `data_version`; retry
open-time validation without excluding a concurrent writer; drop a nested
primary only from mutable cause/context links; preserve both aliases to one
external cause; derive trigger expectations from production-created SQL; or
narrow the evidence claims to the previous partial kill matrix.

Rationale: commit-time rebasing could bless a duplicate-rank semantic split
after either a pending rebuild or a no-rebuild transaction. `IndexStore` now
opens a private SQLite connection subclass with statement caching disabled. Its
non-replaceable dispatcher honors a caller authorizer first and increments a
process-local epoch for each allowed graph, mirror, trigger, or relevant schema
write. Commit publishes the captured epoch verbatim; rollback rebases only
after SQLite restores state. Metadata, diagnostics, and staleness writes do not
advance the graph epoch. Repeated identical SQL through connection execution,
cursor execution, and `executemany` remains observable.

Open-time validation, seven-field and catalog capture, live-seal capture, and
interval-cache warming now share one `BEGIN IMMEDIATE` snapshot, so a WAL or
DELETE writer cannot commit between validation and sealing. Acquisition belongs
to the cleanup boundary: failures before or after native effect release the
transaction or conclusively close the descriptor. Construction refuses an
already-active caller transaction without disturbing it. Deterministic
validation-barrier and writer-first tests prove both commit orderings.

Exception evidence now rebuilds immutable membership only when necessary,
removes a nested primary and later duplicate members, gives the first causal
occurrence ownership of each external identity and complete descendant chain,
and clears later aliases or cycles. Existing no-rebuild group identity remains
unchanged. Graph success/body cleanup, store transaction/close/context cleanup,
and post-commit hook/finalizer composition retain exact primary identity and
each distinct evidence object once.

The permanent catalog suite grows from 52 to 70 cases. It independently freezes
all 18 trigger names, tables, operations, and SQL bodies; covers both catalog
directions; asserts exact post-DML content; proves both physical backends; and
kills nine DDL variants spanning both checks, every `NOT NULL`, composite-key
composition/order, rowid identity, and extra columns. A separate 62-case
mechanics suite covers both backends, all five graph surfaces, managed coherent
restoration before/after rebuild, non-graph controls, rollback, authorizers,
constructor concurrency/acquisition, and the new exception topologies. Fresh
verification passes the complete 919-test P4 slice in 153.74 seconds, the new
132-test mechanics/catalog matrix in 11.34 seconds, and all 1,655 repository
tests in 470.73 seconds at 89.47% core branch coverage. Exact F09b passes in
27.81 seconds; fixture/oracle passes 21 tests in 26.57 seconds; all 14 fixtures
regenerate deterministically; Ruff, formatting, Pyright, lock, diff, sdist, and
wheel checks are clean. The candidate remains gate-pending until stateless
formal reviews inspect one new exact fingerprint.

## 2026-07-28 P4 — Eleventh formal split finds schema shadows and residual proof gaps

Decision: charge global invocation #43 / R-mech #23 and global invocation #44 /
R-test #18 as `REVISE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`6ef43eb612a6213b4f0b4dc6127490f7038a412f`, and cached-diff blob
`6a465e93ec72cc4f72e0f1950ab62aba580af15c` (33 files, 15,890
insertions, 265 deletions). Reopen both formal domains after root-cause
remediation under the unbounded-review authorization.

Alternatives considered: treat TEMP schema objects and `writable_schema` as
unsupported caller behavior; rely on fresh-facade rejection after cached
queries have already succeeded; claim that `sqlite_master` writes always change
`schema_version`; normalize only chained multi-error groups; accept repeated
standalone group members as presentation-only; preserve `__cause__` without its
explicit suppression metadata; let constructor cleanup use a weaker local
composer; infer public query completeness from persisted edge completeness; or
count a denied metadata write as proof of denied graph-write ordering.

Rationale: R-mech #23 demonstrated that exact TEMP copies of
`graph_spatial_state` and `graph_rank_keys` can shadow corrupt persistent main
tables during construction and every cached query. It also enabled
`writable_schema` before facade construction and changed the persisted SQL of a
required trigger without advancing the graph epoch, `data_version`, or
`schema_version`. Both attacks let the cached facade answer while a fresh
facade rejected the same persistent sidecar. All authoritative graph objects
must therefore be explicitly main-qualified, relevant TEMP shadows rejected,
and direct schema-catalog writes included in the process epoch.

The same reviewer found that standalone cleanup groups retain repeated immutable
members because the normalizer mutates only causal links, and rewriting a cause
can flip an explicitly false `__suppress_context__` flag. Constructor cleanup
also chains raw cleanup errors without removing a nested primary, producing
three occurrences and a cause/member cycle. One returning sanitizer must rebuild
membership, preserve group and exception metadata after link rewrites, and be
used at constructor, standalone, and chained boundaries.

R-test #18 independently showed that the fixture I12 test enumerates only hops
already returned by `direct_precedents`, so a selective public-query omission
survives despite exact persisted-edge checks. It also found that authorizer
denial covers only a metadata update, not a denied graph update whose epoch must
remain unchanged, and that constructor before/after-effect acquisition failures
run only under WAL. Remediation must derive expected public hops independently,
freeze denied graph-write ordering across all five surfaces and both backends,
and extend acquisition failure proof to DELETE. Both reviewers verified the
exact fingerprint at entry and exit, ran broad green baselines, and made no
workspace changes; green existing tests do not override the reproduced gaps.

## 2026-07-28 P4 — Schema isolation and returning evidence sanitizer remediated

Decision: bind every authoritative graph read, write, mirror join, state/catalog
lookup, and schema PRAGMA to `main`; reject protected TEMP names and targets at
schema setup, facade construction, and managed mutation entry; and count direct
`sqlite_master`/`sqlite_schema` writes and their TEMP aliases in the private
graph epoch. Retain caller-authorizer precedence so a denied graph update does
not advance trust. Replace the duplicated mutating exception helpers with one
returning sanitizer in `core/exception_evidence.py`, and install its returned
root at query cleanup, constructor cleanup, store close, and post-commit
standalone or chained boundaries.

Alternatives considered: rely only on TEMP-object rejection while leaving
authoritative SQL unqualified; rely only on main qualification while allowing a
TEMP schema to imitate protected state; assume direct schema-catalog writes
always advance SQLite's schema version; mutate immutable group members in place;
normalize only multi-error wrappers; or cover the sanitizer helper without
proving that each caller installs a rebuilt root.

Rationale: main qualification removes name-resolution ambiguity, while explicit
TEMP rejection also covers trigger targets and RTree shadow families and makes
unsupported connection-local interference visible before work begins. The
authorizer observes writable-schema catalog writes even when SQLite leaves
`schema_version` unchanged, but evaluates the client's denial first. The shared
sanitizer derives groups only when membership changes, preserves notes,
traceback, cause, context, and explicit suppression after link assignment, and
gives each recursive membership or causal identity exactly one owner. Constructor
acquisition cleanup uses it for failures both before and after native `BEGIN
IMMEDIATE` effect.

The I12 fixture assertion now starts from the hand-authored frozen edge oracle,
asserts exact `direct_precedents` completeness per source, and round-trips every
independently expected concrete target through `direct_dependents` on both
RTree and interval backends. Authorizer-denial tests reject `UPDATE edges` and
prove relational rows, graph state, catalog, and all five public surfaces remain
unchanged. Writable-schema tests cover connection, cursor, and `executemany`
execution; constructor acquisition tests cover WAL and DELETE; protected TEMP
objects and every standalone exception boundary have both-backend matrices.
Fresh targeted evidence is 123 mechanics cases in 7.44 seconds and 944 graph,
store, fixture, catalog, and mechanics cases in 156.84 seconds. These are
remediation checks, not the final frozen-candidate evidence; a full refresh and
new stateless formal split remain required.

## 2026-07-28 P4 — Non-formal preflight closes full-primary and raw-TEMP gaps

Decision: do not freeze after the first green remediation slice. Incorporate
three non-formal preflight findings: exclude the complete reachable primary
exception graph when preparing chained cleanup, permanently test exception-group
subclass derivation and custom state, and extend the live seal with
`temp.schema_version` so protected TEMP DDL is visible even on a plain
`sqlite3.Connection`. Canonical export now rejects protected TEMP objects and
explicitly binds every projected table to `main`, while retaining its historical
ability to inspect dirty intermediate test/debug state.

Alternatives considered: exclude only the primary root; treat group subclass
preservation as an implementation detail; call the TEMP catalog scanner on every
public query; use `total_changes` as a plain-connection DDL witness; or make
canonical export require a clean graph. The last option broke legitimate P1/P2
intermediate-state and raw-rollback comparisons, demonstrating that TEMP
isolation and graph freshness are separate contracts.

Rationale: a primary group can own an identity through membership, cause, or
context, so excluding only its root still permits one object to appear twice in
the final raised graph. A cycle-guarded traversal now seeds the sanitizer with
every primary-owned identity, and graph, transaction, context-exit, post-commit,
and constructor regressions prove the result. `source.derive()` is now protected
by a custom tagged group test that forces rebuilding and checks subclass, state,
message, members, notes, traceback, and suppression. TEMP DDL does not increment
SQLite `total_changes`, but `PRAGMA temp.schema_version` is a constant-size
connection-local identity; adding it keeps the hot live check bounded and makes
all five raw-connection surfaces fail closed. Canonical export performs its own
protected-shadow check and main-qualified projection rather than borrowing the
stricter graph-clean gate.

The mechanics file now contains 147 cases; its 135-case intermediate result and
the prior section's 123-case result are superseded. Fresh focused checks pass:
364 mechanics-plus-store tests, 157 mechanics-plus-P1 edge cases, and the
RTree-unavailable proxy behavior. The earlier full-suite run occurred while
these preflight edits were still changing and is deliberately rejected as gate
evidence; a stable-tree full refresh remains required before staging.

## 2026-07-28 P4 — Rolled-back TEMP schema reconnect is stabilized before reseal

Decision: absorb SQLite RTree's non-mutating schema-reconnect authorization
replay inside the proven managed-rollback boundary. After rollback, touch both
main RTree virtual tables, then reseal only the process epoch while retaining
the transaction snapshot's `data_version`, main `schema_version`, and TEMP
`schema_version`. Do not weaken the authorizer predicate or ignore RTree shadow
callbacks globally.

Alternatives considered: ignore trigger-origin actions; ignore unqualified
RTree shadow callbacks; make authorizer tracking statement-aware and defer
shadow increments until `total_changes` advances; accept that the first graph
surface works and later surfaces fail; or rebase after the first public query.

Rationale: read-only preflight showed that rolling back TEMP DDL returns the
TEMP schema version and persistent graph state to their exact snapshots, but
the first later RTree connection reload emits 32 write-shaped authorizer
callbacks while changing neither `total_changes`, state, nor catalog. Some
tuples are identical to real direct shadow mutations, so a stateless ignore
rule is unsafe. The managed rollback is the one point where SQLite restoration
has already been proven and no caller code can interleave. Reconnecting both
virtual tables there consumes the replay before capturing the allowed rollback
epoch; preserving the three snapshot version fields prevents cross-connection,
main-DDL, or unresolved TEMP-DDL changes from being blessed. A new both-backend
regression rolls back a protected TEMP table, exercises all five surfaces, and
requires the epoch to remain fixed afterward. The mechanics matrix is now 149
cases and passes in 11.35 seconds; stateless re-preflight and stable-tree broad
verification remain required.

## 2026-07-28 P4 — Stable remediation tree passes the complete author gate

Decision: keep the remediated tree unfrozen through one independent rollback
re-preflight and the complete author verification matrix, then advance it to
candidate-freeze preparation without changing production behavior. The
preflight found no defect in the managed rollback path and made no file change.

Alternatives considered: reuse the earlier full-suite result that ran while the
source was changing; infer RTree mutation safety from successful public reads;
omit deterministic regeneration because fixture hashes were already committed;
or freeze before the branch-coverage suite completed.

Rationale: the exact former RTree reconnect reproduction now leaves all five
public surfaces usable on both backends without epoch movement. Direct mutation
of each backend's spatial table and the base `edges` table still fails closed;
a denied client-authorizer graph update leaves rows, state, catalog, epoch, and
all five surfaces unchanged. Ordinary rollback and protected TEMP-DDL rollback
both passed, and the preflight's focused selection reported 16 passing tests.
The rollback hooks remain internal coordination points that assume SQLite has
already completed rollback; the public `IndexStore` path enforces that order.

Fresh stable-tree verification passes the complete 1,011-test P4 slice in
178.51 seconds, the 219-test mechanics/catalog matrix in 13.92 seconds, and all
1,747 repository tests in 538.73 seconds at 90.05% core branch coverage. The
real F09b guard passes in 28.35 seconds, fixture/oracle verification passes 21
tests in 26.48 seconds, and all 14 generated workbooks reproduce with no diff.
Ruff lint and formatting, Pyright, lock validation, sdist, and wheel builds are
clean. These are author-side results; fresh stateless R-mech and R-test verdicts
must still approve one exact staged fingerprint before the P4 phase gate closes.

## 2026-07-28 P4 — Final preflight closes causal loss and TEMP predicate mutants

Decision: reopen the otherwise green tree for two non-formal adversarial
findings, centralize compound-failure composition, prevent a second sanitizer
pass over already-normalized external members, and exhaustively isolate every
protected TEMP-object predicate before freezing. Do not charge these preflight
loops as formal verdicts.

Alternatives considered: treat lost cleanup evidence as cosmetic; patch only
the constructor despite the same direct composition in successful graph cleanup
and post-commit finalization; retain duplicated boundary-specific prior-cause
logic; accept one `TEMP TABLE edges` example as proof of the complete scanner;
or infer RTree target-prefix coverage from name-prefix coverage.

Rationale: a cleanup root could link to an external exception group whose first
member owned its second member as a cause. The first ownership pass moved the
second member out of group membership, but a later pass revisited the retained
first member and cleared that cause, losing the second exception entirely. A
global normalized-link guard now gives each retained exception one causal pass.
Both the exact topology and a real graph-cleanup boundary retain every distinct
identity once and remain acyclic under a harder shared cycle.

Separately, constructor cleanup installed a new cause without relocating an
existing explicit cause or visible context from its primary failure. One shared
composer now captures that prior evidence, avoids redundant wrapping when it is
already represented, excludes the complete remaining primary graph, sanitizes
the result, and is used by graph, context-exit, transaction, post-commit, and
constructor composition. Both backends prove explicit-cause and unsuppressed-
context preservation at graph-cleanup, post-commit, and constructor boundaries;
acquisition cleanup also carries an existing caused primary.

The original TEMP tests coupled object `name` and `tbl_name`, and coupled RTree
prefix names and targets. The final 246-case mechanics file independently
enumerates all 33 protected names on both backends; uses mixed-case views;
targets unprotected TEMP tables with protected and RTree-prefix trigger names;
targets `main.edges` and both persistent RTree shadow families with unrelated
trigger names; and exercises construction, canonical export, and all five live
surfaces. Fresh read-only mechanics and test rechecks are finding-free.

Final stable-tree verification passes the complete 1,108-test P4 slice in
121.40 seconds, the 316-test mechanics/catalog matrix in 13.82 seconds, and all
1,844 repository tests in 356.45 seconds at 90.05% core branch coverage. Exact
F09b passes in 17.11 seconds; fixture/oracle verification passes 21 tests in
15.84 seconds; all 14 workbooks regenerate without a diff; Ruff, formatting,
Pyright, lock, diff, sdist, and wheel checks are clean. The tree is now ready
for one exact staged fingerprint and fresh formal R-mech/R-test verdicts.

## 2026-07-28 P4 — Twelfth formal split finds two capability-boundary gaps

Decision: charge global invocation #45 / R-mech #24 as `REVISE` and global
invocation #46 / R-test #19 as a clean `APPROVE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`ff164c826168e23ad3af279e6c758f40a9e45c49`, and cached-diff blob
`65c2ef8eac566bf476cd8c0bf8e6df3e06ac1524` (34 files, 17,240
insertions, 319 deletions). Reject the shared candidate and invalidate its test
approval because mechanics remediation will change the tree.

Alternatives considered: treat base-descriptor calls as unsupported despite
the public native connection capability; assume a Python override cannot be
bypassed; rely only on the authorizer epoch after removing `total_changes` from
the seal; let constructor cleanup use one virtual close attempt; mask an
initialization failure with cleanup failure; or mark a connection closed without
proving native descriptor closure.

Rationale: R-mech #24 called `sqlite3.Connection.set_authorizer(connection,
None)` through the base descriptor, bypassing `_TrackedConnection`'s override
and removing the private dispatcher. It then performed the already-frozen
coherent relational/catalog/state restoration on the same handle. The graph
epoch no longer moved, all other seal fields matched, and both backends returned
the forged split. The live seal needs an independent native mutation witness or
the raw capability must be controlled; a permanent both-backend base-descriptor
displacement regression is required.

The same reviewer injected an initialization failure after `BEGIN IMMEDIATE`
and a connection subclass whose virtual `close()` fails before the native
descriptor. `IndexStore.__init__` leaked the transaction and writer lock while
the close error replaced the initialization primary. Constructor cleanup must
prove rollback or conclusive base-descriptor closure and preserve the primary
with sanitized cleanup evidence, including failures before and after
acquisition. R-test #19 independently found no critical, major, or minor test
or oracle defect, reproduced the 1,108-test P4 slice, and verified the exact
fingerprint unchanged; that approval cannot close a candidate rejected in the
mechanics domain.

## 2026-07-28 P4 — Capability membrane and complete constructor ownership

Decision: remediate R-mech #24 by replacing the public native SQLite handle
with a narrow connection/cursor capability membrane and by making every
constructor, configuration, schema-recreation, rollback, and close boundary
prove native cleanup while retaining the exact primary exception.

Alternatives considered: document base-descriptor calls as unsupported while
still returning a native object; retain a raw connection and add another
mutable seal; trust virtual `in_transaction`, `rollback`, or `close` methods;
start constructor ownership only after capability allocation; or rely on the
outer constructor to repair an inner recreation boundary after it had already
replaced the primary failure.

Rationale: the supported public surface now returns one dynamically resolving
connection capability, wraps every cursor and its `.connection`, keeps context
manager entry inside the membrane, rejects custom cursor and row factories,
and chains caller authorizers behind the non-displaceable private dispatcher.
Native SQLite base descriptors reject these facades. The retained capability
follows a legitimately replaced handle and preserves SQLite's closed-handle
error after close without permitting post-close acquisition.

Constructor cleanup now owns the native handle before capability allocation,
uses base-descriptor transaction-state, rollback, and close fallbacks, and
composes unique acyclic cleanup evidence without changing primary identity.
The same rule covers early and final connection configuration steps, actual
schema initialization after `BEGIN IMMEDIATE`, and database recreation. Both
old and temporary replacement handles are conclusively closed. The temporary
file descriptor is cleanup-owned from `mkstemp` acquisition, and pre-/post-
effect descriptor failures cannot leave a `.rebuild` artifact. Post-commit
replacement-close failure leaves the original database intact because
`os.replace` has not occurred.

Permanent regressions cover the five public SQL acquisition paths on both graph
backends, caller authorizer observe/deny/remove behavior, public mutation
invalidation, capability construction failure, early/final configuration
failure, rollback and close before/after native effect, successful-build
post-commit close failure, replacement-schema failure, physical closure,
writer reacquisition, primary/evidence identity, original-database preservation,
and temporary-artifact cleanup.

## 2026-07-28 P4 — Prospective single-reviewer gate protocol

Decision: apply the user's simplified review workflow prospectively without
changing any historical accounting. The orchestrator owns implementation,
debugging, test design, mutation reasoning, documentation, evidence, and all
verification. After one exact candidate is staged with no unstaged or untracked
candidate changes, exactly one independent reviewer returns two explicit
verdicts—one R-mech and one R-test—on that same fingerprint. Each verdict is
charged separately. A revision returns the newly frozen candidate to that same
reviewer through a follow-up task.

Alternatives considered: retroactively merge or delete prior verdicts; continue
non-formal preflight agents; use parallel mechanics and test reviewers; or let
reviewers implement remediation.

Rationale: the new protocol keeps the unbounded quality loop while reducing
coordination overhead and preserving the frozen rule that the ledger counts
verdicts, not subagent invocations. All existing commits, evidence, code, tests,
and global invocations #1–#46 remain unchanged.

## 2026-07-28 P4 — Final author verification after capability remediation

Decision: accept the orchestrator-owned remediation as ready for candidate
freeze after rerunning the complete stable author matrix from the final source
and regenerated fixtures. Formal approval is still pending and must use the
single combined reviewer protocol above.

Rationale: the exact P4 slice passes 1,200 tests in 105.36 seconds; the
mechanics/catalog set passes 330 tests in 13.70 seconds; and the real F09b
boundedness guard passes in 16.60 seconds. The full repository passes all 1,936
tests in 370.51 seconds at 90.05% core branch coverage. Fixture generation and
the parser oracle pass 21 tests in 17.58 seconds, and all 14 generated workbooks
reproduce without a diff. Ruff lint and format, Pyright (zero errors/warnings),
lock validation, sdist, wheel, and diff checks are clean. These results are
author evidence only; the P4 gate remains open until both formal verdicts
approve one unchanged staged fingerprint.

## 2026-07-28 P4 — First combined formal review finds `SQLITE_IGNORE` bypass

Decision: charge global verdict #47 / R-mech #25 and global verdict #48 /
R-test #20 as `REVISE` on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`3f70058ba0ed52714bd842f712e53fe9b6034cab`, and cached-diff blob
`82620eff8209c5ac8c143cda9064bc77c36fdac8` (34 files, 18,631
insertions, 349 deletions). Reject the shared candidate and return its eventual
replacement to the same combined reviewer.

Alternatives considered: treat every non-`SQLITE_OK` caller verdict as a denied
mutation; infer `SQLITE_IGNORE` behavior from `SQLITE_DENY`; or carry the test
approval because the full suite and coverage gate were green.

Rationale: SQLite permits `SQLITE_IGNORE` to continue some operations with
altered semantics. In particular, ignoring `SQLITE_DELETE` disables truncate
optimization but still deletes rows. Returning that caller verdict before the
private dispatcher advanced graph authority allowed an edge deletion while
ignored trigger updates preserved the trusted graph state and private epoch.
Both backends then returned silently changed topology from an already-sealed
graph facade. The reviewer reproduced the defect on the unchanged fingerprint,
reproduced all 1,200 P4 tests and all 1,936 repository tests at 90.05% core
branch coverage, and correctly identified the missing behaviorally distinct
both-backend regression. `SQLITE_DENY` must remain non-mutating; graph-affecting
`SQLITE_IGNORE` must conservatively advance private authority before SQLite
continues.

## 2026-07-28 P4 — `SQLITE_IGNORE` remediation and stable author evidence

Decision: compute graph-authoritative action identity before invoking caller
policy; preserve `SQLITE_DENY` as a non-mutating early return; and advance the
private graph epoch before returning `SQLITE_IGNORE` for a graph-affecting
action. Add one behaviorally distinct both-backend regression rather than a
mechanical action/table matrix.

Alternatives considered: increment for every non-OK verdict, including denied
and invalid verdicts; forbid caller authorizers; depend on `total_changes`; or
test only that the epoch moved without proving SQLite performed the ignored
write and every cached graph surface rejected its old seal.

Rationale: the regression first seals all five graph surfaces, installs a
caller policy that ignores DELETE and UPDATE authorization actions, deletes a
real edge, and proves SQLite removed that row while ignored trigger updates did
not provide the normal state signal. It then requires the private epoch to have
advanced and all five surfaces to return `E_CORRUPT`. The existing denial test
still proves an aborted graph write changes neither rows, state, catalog,
authority, nor public results. Removing the new `SQLITE_IGNORE` branch makes
this regression fail, so the test directly protects the reproduced defect.

Fresh post-remediation verification passes the 1,202-test P4 slice in 119.48
seconds, the 332-test mechanics/catalog set in 12.48 seconds, and all 1,938
repository tests in 365.82 seconds at 90.05% core branch coverage. The real
F09b guard passes in 17.76 seconds; fixture/oracle verification passes 21 tests
in 17.36 seconds; all 14 generated workbooks reproduce without a diff; and
Ruff, formatting, Pyright, lock validation, sdist, wheel, and diff checks are
clean. The replacement candidate must now be staged, fingerprinted, and sent
back to the same combined reviewer.

## 2026-07-28 P4 — Combined follow-up approves the remediated fingerprint

Decision: charge global verdict #49 / R-mech #26 and global verdict #50 /
R-test #21 as clean `APPROVE` verdicts on base
`6eb092b12c1c0398ac76a6153c52096f08904d7a`, staged tree
`fdfa3d0e00faba1844d9636e650b63ee8a395a4c`, and cached-diff blob
`67fb916469f7a109b1f584907c0128f1d7164f3d` (34 files, 18,748
insertions, 349 deletions). Close the P4 formal gate and permit the P4 milestone
commit; do not begin P5 before that commit exists.

Rationale: the same combined reviewer verified the replacement fingerprint at
entry and exit with no unstaged or untracked changes, found no blocking issue,
and independently reproduced the 10-case authorizer selection, 332-case
mechanics/catalog set, and 1,202-case P4 slice. It confirmed that `SQLITE_OK`
tracks normally, `SQLITE_DENY` aborts without advancing authority, and a
graph-affecting `SQLITE_IGNORE` advances authority before SQLite can continue
with altered semantics. The both-backend regression physically deletes the
edge and makes every cached graph surface fail closed, while the separate denial
regression proves the non-mutating contract. Ruff, formatting, and Pyright were
also clean. Both verdicts apply to the same unchanged fingerprint; the only
post-verdict edits are this append-only verdict record, PLAN ledger/checklist
accounting, gate and current-status documentation, and their contract
assertions.

## 2026-07-28 P5 — One catalog, phase-owned producers

Decision: centralize all 13 HANDOFF section 5.6 identities and public records in
`core/diagnostics.py`; preserve the existing P2/P3/P4 producers; persist P5's
cached-error, link-health, and volatile-block findings; and provide typed
constructors for the P6 `I_STALE` and P7 `W_REGEX_TIMEOUT` producers without
persisting placeholder rows.

Alternatives: move every historical producer into one new analyzer (large
regression surface); defer the two later-phase codes entirely (an incomplete
catalog); or manufacture rows without their owning runtime state (dishonest
evidence). The selected boundary keeps proven behavior stable while making
severity, shape, filtering, and future construction uniform.

Rationale: F08 now proves `E_ERRVAL` keys only on OOXML `t="e"`, including an
unrecognized value; F10 resolves a genuine numeric external-link map relative
to the workbook directory without disclosing its target; F11/F18 prove dynamic
and one-per-block volatile findings. The filtered query counts before limiting,
fails closed on corrupt rows, and a diagnostics-analysis version forces old
sidecars through a complete semantic refresh. The full repository currently
passes 1,964 tests in 176.52 seconds; formal review is still pending.

## 2026-07-28 P5 — Main-agent quality loop and freeze readiness

Decision: keep P5 remediation and mutation reasoning in the main agent, add
only serialization-bound and aggregate-integrity regressions, and stop the
author loop once the complete verification matrix was green rather than
expanding into a synthetic combinatorial test grid.

Alternatives: retain the earlier 80% focused diagnostics coverage; add tests
for every operating-system path parser branch; or delegate additional cleanup
and mutation passes. The selected checks cover behavior that can change the
public response: arbitrary OOXML error text is bounded, nested `related` data
is immutable and JSON-shaped, and severity/code aggregate counts cannot claim
an internally inconsistent total.

Rationale: the focused diagnostics module passes 24 tests at 90.43% branch
coverage. The complete repository passes all 1,968 tests in 218.88 seconds and
again under branch instrumentation in 496.61 seconds at 90.11% total core
coverage (`core/diagnostics.py` rounds to 90%). Ruff, format, Pyright, fixture
regeneration, `uv lock --check`, sdist/wheel construction, and `git diff
--check` all pass. Formal review remains pending until this exact candidate is
staged and fingerprinted.

## 2026-07-28 P5 — First combined formal review reproduces four boundary failures

Decision: charge global verdicts #51/#52 as R-mech #27 and R-test #22,
respectively, both `REVISE`, on base
`2024813fc6b34264212a5f89903b5f8391b2030b`, staged tree
`ce38f057ff807a5c753807ad990dfb3784ded1c8`, and cached diff
`8d373ca58622c3d8718d7e3aa6655ef7e4e679fa` (23 files, 1,613 insertions,
33 deletions). The same combined reviewer will receive the remediated frozen
candidate.

Alternatives: treat external-link findings as a workbook-only snapshot; accept
redacted basenames as internal identity; validate only returned rows; or defer
network-path handling to the server. Each would leave a reproduced P5
correctness or security failure, so all findings are in scope for main-agent
remediation.

Rationale: the reviewer independently reproduced a missing-link diagnostic
remaining after the target appeared while workbook bytes stayed unchanged, and
a malformed 101st matching row escaping validation behind the 100-row limit.
It also proved from the code that UNC and authority-bearing file URIs can reach
filesystem I/O, and that distinct numeric M1 links sharing a redacted basename
can acquire the wrong source sheet. The associated test verdict requires exact
toggle, no-network-I/O, numeric-ownership, post-cap corruption, and non-finite
JSON regressions. The reviewer changed no files and verified the frozen
fingerprint at entry and exit.

## 2026-07-28 P5 — External-health and full-filter remediation closes reproduced failures

Decision: persist a path-free external-health snapshot in lifecycle metadata;
recompute external diagnostics and bump generation when that snapshot changes;
classify network forms before filesystem access; recover exact numeric link
ownership by reanalyzing only external-edge source blocks; and stream-validate
every filtered diagnostic before materializing the public page.

Alternatives: force a full workbook reindex for target health; add an external
link index column to the P4 edge schema; or validate a second uncapped SQL query
before the existing capped query. The selected design avoids rewriting
unchanged sheets and the approved graph schema, retains numeric identity from
the already persisted anchor formulas, and validates/counts in one bounded-
memory pass.

Rationale: permanent regressions reproduce both missing→present and
present→missing target transitions with unchanged workbook bytes, three
network-path forms with `Path.is_file` replaced by a trap, two numeric links
whose public basenames collide, corrupt JSON at matching row 101, and all three
non-finite floats. The focused diagnostics module passes 28 tests at 92%
branch coverage; the combined remediated regression set passes 127 tests in
31.53 seconds. The complete repository passes 1,972 tests in 199.86 seconds
and again under branch instrumentation in 386.02 seconds at 90.06% total core
coverage. Final reproducibility checks and the replacement freeze remain.

## 2026-07-28 P5 — Second combined review finds excessive-related-depth escape

Decision: charge global verdicts #53/#54 as R-mech #28 and R-test #23,
respectively, both `REVISE`, on base
`2024813fc6b34264212a5f89903b5f8391b2030b`, staged tree
`7a2fae24cb069931da2c82ccdef3f619b63380d4`, and cached diff
`460332170b554351dc4b258eb7df1f6513f49685` (24 files, 1,936 insertions,
39 deletions).

Alternatives: catch only the observed `RecursionError`; rely on the response cap
to make deep structures harmless; or replace recursive freezing with a new
iterative object builder. The selected remediation combines a documented small
depth limit with defensive recursion shaping at the persisted-data boundary,
preserving the simple immutable representation while preventing traceback
escape even if the JSON decoder itself exhausts recursion first.

Rationale: the reviewer confirmed all five first-review findings were fixed,
then inserted a syntactically valid roughly 2,000-level `related` object beyond
ordinary public expectations. `json.loads` reached recursive freezing, which
leaked `RecursionError` instead of `E_CORRUPT`. The same reviewer will receive a
new frozen fingerprint after a permanent beyond-page nested-payload regression
and broad main-agent verification.

## 2026-07-28 P5 — Related-depth remediation passes the complete author gate

Decision: cap diagnostic `related` containers at 64 nested levels, raise a
typed shape error before recursive construction can exhaust Python, and also
catch `RecursionError` at the persisted row boundary in case JSON decoding
fails first.

Alternatives: only lower Python's global recursion limit; accept arbitrary
depth and implement an iterative freezer; or catch the exception without a
public constructor invariant. The selected local bound is deterministic,
doesn't alter process-global behavior, and keeps both direct and persisted
diagnostics fail closed.

Rationale: the permanent test stores a valid 2,000-level object in matching row
101, proving both full-filter validation and recursion shaping while the public
constructor separately crosses the 64-container limit. The focused diagnostics
and ledger slice passes 41 tests. The complete repository passes 1,973 tests in
203.94 seconds and again under branch instrumentation in 374.51 seconds at
90.01% total core coverage (`core/diagnostics.py` 92%). Final lint, type,
fixture, lock, package, contract, and whitespace checks remain before the new
freeze.

## 2026-07-28 P5 — Combined formal approval closes the phase

Decision: charge global verdicts #55/#56 as R-mech #29 and R-test #24,
respectively, both clean `APPROVE`, on base
`2024813fc6b34264212a5f89903b5f8391b2030b`, staged tree
`bf7e9f2787df2110ae7b59f7aee7b82300dddc88`, and cached diff
`da558fb07f4320cbca1bd8876b35792baca46ce8` (24 files, 2,027 insertions,
39 deletions). Close every P5 checklist item and proceed to its milestone
commit before beginning P6.

Alternatives: carry either prior approval across a changed candidate; ask a
second reviewer to repeat the combined gate; or change implementation after
approval. The exact-fingerprint protocol requires neither: the same combined
reviewer returned both verdicts for the unchanged third tree with no findings.

Rationale: the reviewer verified the base, staged tree, and cached-diff hash at
entry and exit; confirmed all earlier external-health, no-network-I/O, numeric
ownership, post-cap corruption, finite-JSON, and nested-related remediations;
and independently reran the 41-test focused slice, Ruff, formatting, and
Pyright. Both verdicts apply to one unchanged fingerprint. The only
post-verdict edits are this append-only record, ledger/checklist accounting,
verified-status documentation, evidence finalization, and contract assertions.

## 2026-07-28 P6 — Main-agent ownership and interrupted-state reconciliation

Decision: continue P6 entirely in the orchestrator under the user's simplified
review protocol. Do not restart any completed P4/P5 reviewer, do not spawn a
cleanup or mutation agent, and reserve exactly one combined formal reviewer for
the frozen P6 candidate.

Alternatives: revive the earlier non-formal agents, delegate the P6 audit, or
wait for already-completed work. The live agent inventory showed only the root
running; the named P4 preflight and P4/P5 combined reviewers were completed.
The interrupted writer hardening patch had applied, so it was inspected and
continued rather than replayed.

Rationale: this preserves all historical work and accounting while complying
with the immediate workflow amendment. The main agent owns implementation,
tests, mutation reasoning, docs, evidence, and verification through freeze.

## 2026-07-28 P6 — Surgical writer and direct semantic patch

Decision: implement edits as a ZIP-preserving lxml patch followed by a sparse
index transaction. The writer checks the Excel lockfile before source conflict,
orders inserted rows/cells, preserves styles, uses inline strings, expands an
entire shared group before member edits, rejects multi-cell arrays, updates
calculation metadata, deletes the complete calc-chain triple, validates a
same-directory temporary archive, and atomically replaces the source.

Alternatives: an openpyxl load/save round trip; a whole-workbook reindex after
every edit; or in-place XML/ZIP mutation. The first violates I18, the second
discards the specified LSP-like edit path, and the third lacks an atomic source
boundary.

Rationale: touched worksheets are streamed after replacement, but only
requested and expanded shared-formula records are retained. One store
transaction replaces sparse cell rows, recomputes touched-sheet regions and
formula analysis, writes stale rectangles plus `I_STALE`, updates hashes/stat,
and increments generation. If that direct transaction fails after source
replacement, incremental recovery rebuilds the derivable sidecar and records
the precomputed stale set.

## 2026-07-28 P6 — Column fills, I18 evidence, and Excel smoke

Decision: keep the public 500-cell limit on `write_cells`, while allowing
`set_column_formula` to fill the entire resolved column body through a private
validated capacity and one body rectangle for graph propagation. A1 patterns
translate from the anchor; R1C1 patterns render independently at each row.
Reject mixed modes, boundary escapes, excessive Excel UTF-16 text/formula
lengths, invalid Unicode/XML text, and numeric values outside the finite Excel
domain before source replacement.

Alternatives: inherit the bulk-write cap in the semantic fill tool; enumerate a
large body as hundreds of stale seed rectangles; or trust only protected-part
spot checks. Those choices contradict the separate tool contract, add avoidable
graph work, or leave I18 underproved.

Rationale: a 501-body-row ListObject regression protects the distinct capacity.
F16/F21 manifests list every ZIP part's before/after SHA-256 and are reproduced
by a committed renderer; the 50-script Hypothesis property checks untouched
part identity and exact reparsed writes. Desktop Excel 16.0 build 19530 normally
opened a surgically edited F03, recalculated `Summary!C10` to `2232.48`, saved
and closed it, and explicit recalculated refresh cleared the stale state.

## 2026-07-28 P6 — Post-replacement concurrency and final property audit

Decision: compare the parser's post-replacement whole-file hash with the exact
hash installed by the writer before applying a sparse sidecar patch. On a
mismatch, rebuild the sidecar from current workbook bytes, retain the planned
staleness, and return `E_CONFLICT` rather than success. Also require the random
I18 property to prove the complete after-part set equals the before set minus
declared deletions and that the actual byte-difference set equals the writer's
declared modifications.

Alternatives: let a later freshness call discover the mismatch; return success
after index recovery; or compare only pre-existing untouched payloads in the
property. Each alternative can hide a second writer, an undeclared added ZIP
member, or an incomplete modification manifest.

Rationale: the permanent race wraps the real atomic writer, changes a second
worksheet before direct index application, proves `E_CONFLICT`, and proves the
recovered index contains the second writer's value. The final semantic-fill
audit also adds a native ListObject totals-row case: four body formulas change
while the `SUBTOTAL` totals formula remains untouched. The branch-instrumented
repository run passes 2,020 tests at 89.49% total core coverage; the live test
remains deliberately deselected from automated runs.

## 2026-07-28 P6 — Pre-freeze verification complete

Decision: freeze only after one final uninstrumented repository run and a clean
desktop-Excel rerun following COM ownership cleanup.

Alternatives: rely on the earlier 2,019-test run before the totals-row
regression, treat the branch-instrumented run as the only broad result, or keep
the passing live assertion despite its teardown warning. Each would leave the
final candidate or live harness less precisely verified than the staged tree.

Rationale: the final repository passes 2,020 tests with one live test
deliberately deselected in 194.34 seconds; the same 2,020 pass under branch
coverage in 380.71 seconds at 89.49% core coverage. The P6-focused slice passes
73 tests in 22.39 seconds. Excel 16.0 build 19530 passes the live smoke in 3.21
seconds after releasing every COM proxy before apartment teardown. Fixture
regeneration emits all 20 current IDs without tracked drift; lock validation,
sdist/wheel build, Ruff, format, Pyright, and whitespace checks pass.

## 2026-07-28 P6 — Combined formal review #1 requires schema-order remediation

Decision: charge global verdicts #57/#58 as R-mech #30 and R-test #25,
respectively, both `REVISE`, on base
`672e91ad362dcf1d984307fe4c2fd9db89d72b04`, staged tree
`fd797eb9796ad561899c8ab9439e44fbd45619a4`, and cached diff
`17eba11d65fbde3e64854170c2cb9e4f4b6ca8e2` (34 files, 3,883 insertions,
82 deletions). Keep implementation ownership in the main agent and return the
new frozen candidate to this same reviewer.

Alternatives: dismiss the findings because the ordinary F03 live smoke passed;
add only tests; or ask a second reviewer. The exact OOXML sequences permit
valid trailing cell extensions and several workbook children after `calcPr`,
so ordinary generated fixtures did not exercise the defect. The user's review
protocol requires main-agent remediation and the same combined reviewer.

Rationale: replacement `<f>`, `<v>`, or `<is>` content was appended after a
preserved cell `extLst`, and a missing `calcPr` was inserted only before
workbook `extLst`. The fix inserts cell value content before its extension list
and creates `calcPr` before every schema successor from `oleSize` through
`extLst`. One unit workbook covers inline string, numeric value, and formula
cells with extensions; another removes `calcPr` and adds `fileRecoveryPr`.
Desktop Excel normally opens, recalculates, saves, and closes that adversarial
combination, giving the repair-sensitive ordering a live canary as well.

## 2026-07-28 P6 — Remediated candidate verification refresh

Decision: freeze only after rerunning the entire verification matrix on the
schema-order remediation. Expand the workbook-successor regression across all
nine valid children after `calcPr`, rather than proving only the live
`fileRecoveryPr` case.

Alternatives: retain the pre-review broad results because the remediation was
localized; or keep a single successor regression. Both would leave the formal
review without fresh whole-repository evidence, and a missing name in the
ordered-successor set can independently change output validity.

Rationale: the P6-focused editor/R1C1/property/fixture/oracle slice passes all
83 cases in 26.31 seconds, including 45 editor-unit cases in 5.19 seconds. The
normal suite passes 2,030 tests with one opt-in live test deselected in 242.17
seconds. The branch-instrumented suite passes the same 2,030 tests in 438.27
seconds at 89.58% total core coverage. Desktop Excel 16.0 build 19530 passes
the adversarial extension/order smoke in 2.75 seconds with clean COM teardown.
Fixture regeneration emits all 20 IDs with no tracked drift; the 70-package
lock check, sdist/wheel build, Ruff, format, Pyright, README contract, and
whitespace checks pass.

## 2026-07-28 P6 — Combined formal review #2 requires retry-race remediation

Decision: charge global verdicts #59/#60 as R-mech #31 and R-test #26,
respectively, both `REVISE`, on base
`672e91ad362dcf1d984307fe4c2fd9db89d72b04`, staged tree
`5671fbba8cf84eed0cf4a1641908e47fd29832dd`, and cached diff
`845e1fd8a8973c718611c802482c0543c980aac9` (34 files, 4,138 insertions,
86 deletions). Keep all remediation in the main agent and return the next
frozen fingerprint to the same reviewer.

Alternatives: rely on the service's post-install hash reconciliation; retry
without delay; or remove the Windows permission retry. Post-install checking
cannot recover bytes already overwritten by the editor, while removing the
bounded retry would discard an intentional cross-platform availability path.

Rationale: if the first atomic replacement raised `PermissionError`, a writer
or Excel session could change the destination during the half-second delay.
The second attempt then overwrote that state without repeating the ordered
lockfile and hash preconditions. The replacement boundary now checks the Excel
lockfile first and compares the current destination to the indexed hash before
every attempt. Four deterministic regressions prove that external bytes survive
as `E_CONFLICT`, an intervening lockfile wins as `E_OPEN_IN_EXCEL`, unchanged
bytes permit retry success, and two permission failures return `E_LOCKED`
without changing the workbook.

## 2026-07-28 P6 — Retry-race candidate verification refresh

Decision: repeat every broad and external verification gate after moving the
lockfile/hash preconditions into the atomic-replacement retry boundary.

Alternatives: run only the four new tests because the change is localized, or
reuse the schema-order candidate's live and packaging evidence. The finding is
a data-loss race at the production mutation boundary, so the next formal
fingerprint warrants independent fresh evidence across all gates.

Rationale: the editor unit module passes 49 tests in 4.95 seconds, and the
editor/R1C1/property/fixture/oracle slice passes 87 in 24.19 seconds. The full
repository passes 2,034 tests with one opt-in live test deselected in 229.95
seconds; branch instrumentation passes the same 2,034 in 422.77 seconds at
89.68% total core coverage. Desktop Excel 16.0 build 19530 passes the
adversarial schema-order smoke in 2.33 seconds with clean COM teardown. All 20
fixture IDs regenerate without tracked drift; the 70-package lock check,
sdist/wheel build, Ruff, format, Pyright, README contract, and whitespace
checks pass.

## 2026-07-28 P6 — Combined formal review #4 closes the phase gate

Decision: charge global verdicts #63/#64 as R-mech #33 and R-test #28,
respectively, both clean `APPROVE`, on base
`672e91ad362dcf1d984307fe4c2fd9db89d72b04`, staged tree
`8c794e703c0d0a783f23095c89d920d29f3c8f10`, and cached diff
`6587b6161e4c15cde11dfcd10077d1bd323ad266` (34 files, 4,540 insertions,
86 deletions). Close P6 only after recording the verdicts and committing the
milestone.

Alternatives: carry only one approval because one reviewer produced both
verdicts, or treat the post-review accounting update as a new implementation
candidate. The user-defined protocol counts verdicts, so the combined review
adds one to each domain; post-verdict ledger/status promotion records the gate
without changing the approved implementation or test mechanics.

Rationale: the reviewer verified the exact fingerprint unchanged at entry and
exit, independently passed all 50 editor tests, and reported no mechanics or
test findings. It explicitly confirmed the stable sidecar generation boundary,
atomic-retry preconditions, OOXML schema ordering, valid second-writer recovery,
part-preservation evidence, property proof, full suites, branch coverage,
desktop-Excel smoke, deterministic fixtures, lint, format, types, build, and
lock checks. P6 is verified and ready for its milestone commit.

## 2026-07-28 P6 — Combined formal review #3 requires sidecar snapshot remediation

Decision: charge global verdicts #61/#62 as R-mech #32 and R-test #27,
respectively, both `REVISE`, on base
`672e91ad362dcf1d984307fe4c2fd9db89d72b04`, staged tree
`2f0755b91417deccb299a33662d8b11de8663106`, and cached diff
`861802ee5ce0b477df1834409cb8a5f59a779d17` (34 files, 4,355 insertions,
86 deletions). Preserve the prior approved remediations and return the next
frozen candidate to the same combined reviewer.

Alternatives: rely on the service's initial installed-hash comparison; persist
the parser's stat rather than the path's later stat; or let ordinary freshness
repair any mismatch. The parser can remain attached to an old file generation
after a path replacement, and a sidecar containing old cells plus the new
path's accepted stat can incorrectly pass freshness.

Rationale: stable SHA-256 and `(device, inode, size, mtime_ns, ctime_ns)`
snapshots now bracket touched-sheet collection, direct index application, and
the transaction commit. Package hashes, parsed cells, and committed mtime/size
therefore describe the same installed bytes. If collection itself fails after
the path changes, the snapshot guard converts that failure to `E_CONFLICT` so
the established recovery path rebuilds the current workbook. The deterministic
race installs a valid workbook with `Calc!A2 = 9999` during collection and
proves conflict return, byte preservation, recovery of that exact value, and a
subsequent no-change refresh.

## 2026-07-28 P6 — Sidecar-snapshot candidate verification refresh

Decision: rerun every broad and external verification gate after introducing
the generation-stability snapshots and mid-collection conflict recovery.

Alternatives: reuse the retry-race candidate's broad evidence or run only the
new race test. The remediation affects the production direct-index transaction
and recovery boundary, so an exact formal fingerprint needs fresh repository,
coverage, live, and packaging evidence.

Rationale: the editor unit module passes 50 tests in 4.04 seconds, and the
editor/R1C1/property/fixture/oracle slice passes 88 in 24.74 seconds. The full
repository passes 2,035 tests with one opt-in live test deselected in 228.31
seconds; branch instrumentation passes the same 2,035 in 426.04 seconds at
89.65% total core coverage. Desktop Excel 16.0 build 19530 passes the
adversarial schema-order smoke in 2.15 seconds with clean COM teardown. All 20
fixture IDs regenerate without tracked drift; the 70-package lock check,
sdist/wheel build, Ruff, format, Pyright, README contract, and whitespace
checks pass.

## 2026-07-28 P7 — Shared MCP/CLI boundary and bounded-query decisions

Decision: implement one stateless `ToolService` beneath exactly 14 FastMCP
handlers and the Typer debugging commands. Bind every request to the existing
freshness lifecycle, use Pydantic-derived schemas, preserve canonical core
errors, and sanitize unexpected failures before they reach the client.

Alternatives: duplicate query behavior in MCP and CLI; expose the SQLite
capability directly from handlers; or let framework exceptions become the
public error surface. A shared service keeps path checks, generation handling,
caps, symbol resolution, and write semantics identical across transports.

Rationale: T6 drives the production stdio process with the official MCP client,
compares all generated schemas and annotations, exercises every happy and error
path, and separately verifies the CLI projections. The `in` field requires a
Pydantic validation alias plus a distinct serialization alias: an ordinary
alias generated the correct JSON schema but caused FastMCP to call Python with
the reserved keyword `in`.

## 2026-07-28 P7 — Real progress, regex timeout, and response-cap mechanics

Decision: add an optional synchronous per-sheet callback to the core lifecycle,
bridge it back to the MCP event loop only when a client progress token exists,
and fall back to reporting the indexed catalog on no-op opens. Add the locked
timeout-capable `regex` runtime dependency and enforce a single two-second
budget across 1,000-character subjects. Measure response size using the pretty
JSON representation FastMCP actually emits.

Alternatives: send progress only after indexing; statically reject patterns
that look risky; use Python `re` with deadline checks only between cells; or cap
compact JSON while ignoring FastMCP whitespace. Post-hoc progress does not help
long cold opens, static regex heuristics produce false positives, and Python
`re` cannot interrupt one catastrophic match on Windows.

Rationale: regressions prove progress originates inside a cold three-sheet
index; `^(a|aa)+b$` against a 1,000-character nonmatch is interrupted with
`W_REGEX_TIMEOUT`; F20 and a 500-edit result remain within 8,000 emitted
characters; a 181-column region omits samples before the 200-value ceiling; and
row-major range pages remain unambiguous when a page crosses a row boundary.

## 2026-07-28 P7 — Cursor, confinement, and phase-boundary decisions

Decision: encode cursors as opaque URL-safe base64 objects containing tool,
parameter hash, offset, and generation. Resolve every path and each configured
`EXCEL_LSP_ROOT` directory before case-normalized containment checks. Register
the `bench` CLI command now but delegate its executable harness and measurements
to P8, which the authoritative phase table assigns to benchmark implementation.

Alternatives: continue pagination after a refresh; compare lexical path
prefixes; allow file-valued roots; or implement the six-task benchmark early in
P7. Generation mismatch must fail with `E_STALE_CURSOR`; lexical prefixes and
unresolved symlinks are not confinement; and advancing P8 work would violate
the ordered phase gates.

Rationale: stdio conformance advances a cursor, performs a surgical write, and
observes `E_STALE_CURSOR`; an external workbook change causes the next read
tool to report `reindexed=true`; unit and subprocess tests cover allowed paths,
outside denial, and symlink escape denial. P7 therefore proves the command and
transport boundary without claiming P8 benchmark results.

## 2026-07-28 P7 — Exact-candidate author verification

Decision: treat `Candidate Pn` as an explicit claims-ledger state between
implementation evidence and formal approval, rerun the complete candidate
after that accounting correction, and freeze only after every broad,
reproducibility, and packaging check describes the same source tree.

Alternatives: label unreviewed P7 claims `Verified`; leave them `Planned`
despite committed candidate evidence; or reuse the earlier 2,049-test run after
adding the maximum-write regression. Each alternative would misstate either
review status or exact-candidate verification.

Rationale: all 12 README contract tests pass with the explicit candidate state;
the complete repository passes 2,050 tests with one opt-in live test deselected
in 206.35 seconds; branch instrumentation passes the same 2,050 tests in 375.15
seconds at 89.63% total core coverage. Ruff lint and format, Pyright, the
70-package lock check, deterministic regeneration of all 20 fixture IDs,
sdist/wheel construction, and whitespace validation are clean. These are
author-side results only; one combined reviewer must still return both formal
verdicts on the frozen fingerprint.

## 2026-07-28 P7 — Combined formal review #1 requires boundary remediation

Decision: charge global verdicts #65–#66, R-mech #34 and R-test #29, as
`REVISE`; reopen the P7 candidate; and remediate every finding in the
orchestrator before sending one replacement fingerprint back to the same
combined reviewer.

Alternatives: treat the response cap as success-only; retain the 1,001-cell
internal cutoff as an undocumented search horizon; infer value-search freshness
from a prior map; define profile count as persisted rows; or add only direct
unit reproductions. Each would leave an agent-visible false claim or an
unverified transport boundary.

Rationale: the reviewer independently reproduced a 12,100-character canonical
error, a missing row-1,100 symbol, a stale cached match without a staleness
flag, and a sparse `A1:A5` profile reporting count two. Expected errors now pass
through the transport fitter; symbol search streams the full domain with a
bounded deterministic top 100 and exact total; value searches expose
scope-conservative staleness; and profile count measures resolved rectangle
positions separately from non-null values. Four focused regressions plus the
real stdio test cover the behaviorally distinct failures. The remediated
repository passes 2,054 tests with one live test deselected in 377.61 seconds;
branch instrumentation passes the same set in 778.73 seconds at 89.67% total
core coverage. All 20 fixture IDs regenerate without tracked drift; Ruff,
format, Pyright, the 70-package lock check, sdist/wheel build, README contracts,
and whitespace checks are clean. Fresh formal approval remains required.

## 2026-07-28 P7 — Combined formal review #2 finds escaped-value expansion

Decision: charge global verdicts #67–#68, R-mech #35 and R-test #30, as
`REVISE` on base `701b6832b5138010aab8063b9dade7371154cb73`, tree
`d6381cd0ef71a16224a952111b380c51d0776719`, and staged diff
`401a17ea09b18525dcb256fbd86a9541379c0aef`. Reopen P7 and keep all
remediation, verification, and regression work in the orchestrator before a
follow-up to the same combined reviewer.

Alternatives: rely on the transport-wide string fitter; define the 4,000-code
point cell guard as sufficient; mark the response as a truncated page; or omit
the value-level signal. Those choices would conflate serialization expansion
with pagination and could silently change a complete one-cell result.

Rationale: the reviewer reproduced a cell containing 4,000 newlines whose
direct `read_range` result occupied 8,243 pretty-JSON characters. FastMCP's
generic envelope fitting shortened the nested value but exposed only top-level
`truncated=true`, no cursor, and no `valueTruncated` signal. The service now
binary-searches the largest deterministic string prefix that fits the actual
pretty-serialized response, marks `valueTruncated=true`, and preserves the
complete page as `truncated=false` with `cursor=null`. A direct service
regression and the real stdio conformance flow exercise this behaviorally
distinct escaped-string boundary. Fresh author verification and a replacement
fingerprint remain required before re-review.

## 2026-07-28 P7 — Escaped-value remediation author verification

Decision: freeze the replacement candidate only after direct and subprocess
escaped-value regressions, a 34-test P7/accounting slice, the complete suite,
branch instrumentation, deterministic fixture regeneration, static checks,
and package construction all pass on the remediated implementation.

Alternatives: rely on the generic envelope fitter's existing tests; skip the
real stdio reproduction; or omit the final exact-accounting suite because only
evidence Markdown changed after the first broad pass. Each would weaken the
same-boundary proof requested by the combined reviewer.

Rationale: the focused slice passes 34 tests; an uninterrupted full run passes
2,055 tests with one opt-in live test deselected in 400.05 seconds; branch
instrumentation passes the same 2,055 tests at 89.67% total core coverage in
785.74 seconds. After final accounting and evidence edits, an exact-candidate
full rerun again passes all 2,055 tests; its 2,007.08-second wall time includes
a prolonged runner-output stall and is retained here rather than substituted
for the uninterrupted measurement. Ruff lint and formatting, Pyright, the
70-package lock check, all 20 deterministic fixture IDs, sdist/wheel build,
README contracts, and whitespace validation are clean. The candidate is ready
for one frozen fingerprint and follow-up to the same combined reviewer.

## 2026-07-28 P7 — Combined formal approval closes the phase

Decision: charge global verdicts #69–#70 as R-mech #36 and R-test #31, both
clean `APPROVE`, on base `701b6832b5138010aab8063b9dade7371154cb73`, tree
`2e7d105387e44a4acbdb674f6b73121e1d2f0d2f`, and staged binary-diff hash
`b3cb3077e40963b328d82fc4fc7095ab1a7a5798` (27 files, 3,127 insertions,
180 deletions). Close the P7 formal gate and permit the milestone commit after
only ledger, phase-status, and approval-evidence promotion edits.

Alternatives: carry either rejected verdict forward; treat the generic
transport fitter as sufficient proof; start P8 before recording the gate; or
request another reviewer after a finding-free combined result. Each would
violate the single-reviewer protocol or obscure the exact approved tree.

Rationale: the same reviewer revisited both prior review rounds and found no
critical, major, or minor issue in either domain. Independent verification
passed the 34-test P7/CLI/accounting slice, targeted Ruff and format checks,
and Pyright with zero findings. A separate two-page JSON-expansion probe
produced 7,999-character pages with correct value truncation and cursor
semantics, while a complete one-cell request produced exactly 8,000 characters
with `valueTruncated=true`, `truncated=false`, and `cursor=null`. Entry and exit
fingerprints matched; there were no unstaged or untracked changes. Both formal
verdicts therefore approve the same frozen implementation, test, and evidence
candidate.
