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
