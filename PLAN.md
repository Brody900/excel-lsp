# Excel LSP delivery plan

Authoritative specification: `HANDOFF.md` (v1.1). Work phases execute in order; a phase closes only after its required gate approves.

## Phase checklist

### P0 — Recon and scaffold

- [x] Record OS, Excel availability/launch, `uv`, git, Python, and Codex CLI/auth.
- [x] Verify SQLite R*Tree support.
- [x] Verify the PyPI distribution name.
- [x] Pin openpyxl and tiktoken; commit `uv.lock`; record the choices.
- [x] Create the package/test/docs/benchmark skeleton and CI stub.
- [x] Create the deterministic fixture-generator skeleton.
- [x] Author and extract the F16 `vbaProject.bin`; document provenance.
- [x] Complete Phase 0 self-check and milestone commit.

### P1 — Parser, storage, and lifecycle

- [x] Implement the single-pass lxml OOXML parser and normalized cell stream.
- [x] Implement package/part hashing, workbook metadata, and supported-format errors.
- [x] Implement SQLite schema, WAL settings, R*Tree/fallback abstraction, and canonical exports.
- [x] Implement full/incremental indexing and freshness lifecycle.
- [x] Generate F01 and make the openpyxl oracle harness green.
- [x] Record openpyxl read-only behavior for shared formulas, tables, and merged cells.
- [x] Pass R-mech and R-test gates.

### P2 — Regions and workbook map

- [x] Implement ListObject-first region detection, headers, merged headers, dtypes, and confidence.
- [x] Implement stable symbols and the compact workbook map.
- [x] Generate/exercise F20 and enforce deterministic degradation plus the 8,000-character cap.
- [x] Meet the F03 1,500-token map budget.
- [x] Create README skeleton and claims-to-artifacts matrix.
- [x] Pass R-test and the user-authorized early R-repo gate; every prior
  invocation remains charged.
- [x] Obtain a fresh R-mech approval after the latest `REVISE`.

### P3 — Formula references and blocks

- [x] Implement reference classification, names, structured/3-D/external refs, and dynamic/volatile flags.
- [x] Implement LET/LAMBDA suppression, modern prefix normalization, spills, and implicit intersection.
- [x] Implement R1C1 normalization, block construction/extrusion/clamping, and inconsistency detection.
- [x] Generate/exercise F19 and verify invariant I20.
- [x] Pass R-mech and R-test gates.

### P4 — Graph and traces

- [x] Implement EdgeStore range queries and graph construction.
- [x] Implement precedents, dependents, path queries, pagination/truncation, and depth caps.
- [x] Implement bounded two-stage circular detection.
- [x] Verify exact graph behavior on F03/F04/F05/F15/F19 and circular behavior on F09a/F09b.
- [x] Pass R-mech and R-test gates.

### P5 — Diagnostics

- [x] Implement every diagnostic code and filters from the handoff.
- [x] Verify cached-error, broken-link, inconsistency, dynamic, volatile, and large-sheet behavior.
- [x] Pass R-mech and R-test gates.

### P6 — Surgical editor and staleness

- [ ] Implement ordered OOXML cell patching and atomic zip replacement.
- [ ] Implement inline strings, formula/value mechanics, shared-formula expansion, and array refusal.
- [ ] Implement calc metadata edits, calcChain triple-delete, lock/conflict preconditions, and dimensions.
- [ ] Implement direct index patching, generation bumps, and transitive staleness.
- [ ] Prove I18 on F16/F21 and with property tests.
- [ ] Complete a live-Excel round-trip smoke test.
- [ ] Pass R-mech and R-test gates.

### P7 — MCP server and CLI

- [ ] Implement all 14 tools with Pydantic schemas and canonical structured errors.
- [ ] Enforce path confinement, regex deadline, freshness, value/character caps, and generation cursors.
- [ ] Add MCP annotations, initialization instructions, and indexing progress.
- [ ] Implement the full CLI including Mermaid graph output and benchmark command.
- [ ] Pass T6 conformance plus R-mech and R-test gates.

### P8 — Live evidence and benchmarks

- [ ] Author and index L1–L3 in desktop Excel.
- [ ] Complete every live protocol step, including VBA run, write refusal, chart preservation, screenshots, and GIF.
- [ ] Implement naive baseline, six task checkers, scripted replays, and headless LLM eval harness.
- [ ] Run both repetitions, enforce the cost guard, preserve raw results, and grade exact answers.
- [ ] Generate five charts, accuracy table, and audit-cost callout.
- [ ] Verify S1–S6 with linked evidence.
- [ ] Pass R-test gate.

### P9 — Documentation and release

- [ ] Complete README in the required order with fair, sourced comparisons and artifact-backed claims.
- [ ] Complete architecture, internals, security, contribution, tool-reference, changelog, examples, and registry copy.
- [ ] Run clean-environment install, CLI, MCP, `uvx`, lint, type, test, coverage, and build checks.
- [ ] Verify the GitHub Actions 3 OS x 3 Python matrix.
- [ ] Create the public GitHub repository and push `main`.
- [ ] Tag/release `v0.1.0`, attach release assets, and publish to PyPI or document the git+ fallback.
- [ ] Prepare official MCP Registry, Smithery, mcp.so, and PulseMCP submissions.
- [ ] Spend final R-repo reviews and close the release gate.

## Review ledger

Original allocation: 10 R-mech, 10 R-test, 10 R-repo. The first amendment
pooled that nominal 30-invocation budget. On 2026-07-27, the user explicitly
superseded the hard ceiling with unbounded fresh-review loops and continued
subagent orchestration. The original allocation and every verdict remain
visible for audit; extra invocations are charged sequentially and never waive
a phase gate. Phase gates may pass on APPROVE with minor findings, but the
release aims for a clean APPROVE in every domain.

| Domain | Used | Nominal remaining | Latest verdict | Notes |
|---|---:|---:|---|---|
| R-mech | 29 | 0 nominal; unbounded retries authorized | APPROVE | P1 used three REVISE verdicts, then a finding-free APPROVE. P2 reviews #1–#4 found and remediated map/configuration and region-scaling majors; #5 approved. P3 #10 and #11 reopened two candidates; #12 approved. P4 #13/#14 found mirror/state/schema gaps; #15 found failed-commit/live-storage errors; #16 found an incomplete live trust seal; #17 found unsnapshotted cross-connection retrieval; #18 found snapshot-release lock leakage; #19 found unprotected acquisition, native-store close, and cyclic causality; #20 found virtual closure-state spoofing and aliased evidence; #21 found recursively nested evidence aliasing and a cached duplicate-rank semantic split; #22 found three unnormalized cleanup compositions and a live cached-facade seal gap; #23 found TEMP shadowing, writable-schema epoch bypass, incomplete standalone group normalization, and constructor-cleanup causal aliasing; #24 found base-descriptor authorizer displacement and incomplete `IndexStore` constructor cleanup; #25 found caller-authorizer `SQLITE_IGNORE` could allow graph mutation without advancing private authority; #26 cleanly approved the remediated frozen candidate. P5 #27 found stale unchanged-workbook link health, network-path filesystem access, redacted-basename ownership collisions, post-cap corruption validation, and non-finite JSON acceptance; #28 reproduced an uncaught recursion failure on excessively nested persisted related data; #29 cleanly approved the fully remediated frozen candidate. |
| R-test | 24 | 0 nominal; unbounded retries authorized | APPROVE | P1 used one REVISE and one APPROVE-with-minor. P2 used one REVISE and one clean APPROVE. P3 #5 and #6 approved superseded trees; #7 approved. P4 #8/#9 found incomplete edge oracles; #10/#12/#13/#15 approved candidates later rejected by mechanics; #11 found incomplete F09b traversal proof; #14 reproduced the native-subclass store-close lock leak; #16 cleanly approved a candidate later rejected by mechanics; #17 found missing permanent kill-proof for exact rank-catalog DDL, content, trigger validation, and current-sidecar reconstruction; #18 found a selectively omissible fixture round-trip oracle, missing graph-write authorizer-denial proof, and WAL-only constructor acquisition-failure coverage; #19 cleanly approved a candidate rejected by R-mech #24; #20 found no both-backend `SQLITE_IGNORE` mutation/invalidation regression; #21 cleanly approved the remediated frozen candidate. P5 #22 found missing regressions for unchanged-workbook link health, no-I/O network targets, numeric same-basename ownership, post-cap corruption, and non-finite JSON rejection; #23 found no excessive-nesting corruption regression beyond the materialized page; #24 cleanly approved the fully remediated frozen candidate. |
| R-repo | 3 | 7 | APPROVE | The first two P2 invocations remain charged as `REVISE`. Fresh review #3 returned `APPROVE` with one architecture-status minor: it still implied that R-test and all three P2 approvals were pending. This current-state update corrects that wording. Reserve at least 3 R-repo invocations for P9. |

Total formal verdicts recorded: 56. Against the historical
allocation, 0 nominal slots remain. At least 3 future R-repo invocations remain
reserved for P9 under the later unbounded-review authorization. Additional fresh
invocations are likewise authorized whenever a `REVISE` or changed candidate
requires them.

### User-authorized review-governance amendment

On 2026-07-16, the user explicitly authorized the smallest amendment needed to
resolve the remaining-gate arithmetic:

- Keep the maximum at 30 review invocations overall.
- Pool otherwise-unused R-mech, R-test, and R-repo slots.
- Permit the required P2 R-repo retry sequence after the first `REVISE`,
  including a fresh re-review after any subsequent `REVISE`, within the
  unchanged overall pool.
- Preserve every phase gate, a fresh stateless reviewer for every invocation,
  and all prior verdict accounting.
- Reserve at least 3 R-repo invocations for P9.

The minimum completion path is now 30 total invocations, leaving no pooled
contingency slots. This amendment changes allocation only; it does not waive or
combine any review, acceptance criterion, phase ordering rule, or release gate.

### 2026-07-27 unbounded-review amendment

After P3's initial formal split, the user explicitly directed Codex to keep
unbounded review loops and subagent orchestration even at high reasoning. This
later instruction supersedes the first amendment's 30-invocation hard stop.
Every invocation remains stateless, fingerprint-bound, and charged in this
ledger; a changed candidate always receives fresh reviews. No approval is
carried across a code change, no gate is combined, and phase ordering remains
unchanged.

### 2026-07-28 single-reviewer amendment

Effective with the next P4 candidate freeze, the user moved all implementation,
debugging, test design, mutation reasoning, documentation, evidence refresh,
formatting, typing, and verification back to the orchestrator. Non-formal
preflight agents and parallel gate reviewers are no longer used. Historical
reviews and their accounting remain unchanged.

For each newly frozen candidate, exactly one independent reviewer evaluates
both domains on the same fingerprint and returns two explicit verdicts:
`R-mech: APPROVE|REVISE` and `R-test: APPROVE|REVISE`. The ledger still charges
one verdict in each domain, so a combined review adds two to the total verdict
count. If either verdict revises, the orchestrator alone remediates and verifies
the new candidate, then returns its new fingerprint to that same reviewer by
follow-up. The unbounded quality loop, exact-fingerprint rule, phase ordering,
and requirement that both verdicts approve one unchanged tree remain intact.

## Definition of Done

- [ ] S1–S7 all verified with evidence linked (S4 includes the part-diff proof)
- [ ] CI green: 3 OS × 3 Python, lint + types + tests, coverage ≥ 85% core
- [ ] All **21** generated fixtures + 3 live-authored fixtures exercised; **T-oracle green with documented skip-list**
- [ ] Live-Excel evidence committed: authoring, round-trip (incl. .xlsm/VBA with macro-runs assertion), trace cross-validation, no-repair screenshots, **open-in-Excel refusal (step 5)**, **chart-intact screenshot (step 6)**, demo GIF
- [ ] Benchmarks: scripted + LLM-eval results committed with raw data, **ANSWER-contract checkers**, **both repetitions reported**, and 5 charts (incl. incremental series); hero chart in README
- [ ] README complete per §10.2 (**14 tools, LSP-style qualifier, Security & scope note**); fresh-env install verified
- [ ] **SECURITY.md, CLAUDE.md, uv.lock committed; `tests/fixtures/README.md` documents vbaProject.bin provenance**
- [ ] Review ledger: budget accounting complete (incl. the P2 R-repo spend); every domain ended on clean APPROVE (or exhaustion policy executed)
- [ ] `v0.1.0` tagged; PyPI published or fallback documented; registry submissions prepared
- [ ] agent-log.md tells the whole story; final summary written for the user listing follow-ups (tokens, registry accounts)
