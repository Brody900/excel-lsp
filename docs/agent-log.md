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
