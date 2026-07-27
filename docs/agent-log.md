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
