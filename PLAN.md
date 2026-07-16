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

- [ ] Implement reference classification, names, structured/3-D/external refs, and dynamic/volatile flags.
- [ ] Implement LET/LAMBDA suppression, modern prefix normalization, spills, and implicit intersection.
- [ ] Implement R1C1 normalization, block construction/extrusion/clamping, and inconsistency detection.
- [ ] Generate/exercise F19 and verify invariant I20.
- [ ] Pass R-mech and R-test gates.

### P4 — Graph and traces

- [ ] Implement EdgeStore range queries and graph construction.
- [ ] Implement precedents, dependents, path queries, pagination/truncation, and depth caps.
- [ ] Implement bounded two-stage circular detection.
- [ ] Verify exact graph behavior on F03/F04/F05/F15/F19 and circular behavior on F09a/F09b.
- [ ] Pass R-mech and R-test gates.

### P5 — Diagnostics

- [ ] Implement every diagnostic code and filters from the handoff.
- [ ] Verify cached-error, broken-link, inconsistency, dynamic, volatile, and large-sheet behavior.
- [ ] Pass R-mech and R-test gates.

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

Original allocation: 10 R-mech, 10 R-test, 10 R-repo. The
user-authorized amendment below keeps the frozen 30-invocation overall ceiling
while pooling otherwise-unused domain slots. Phase gates may pass on APPROVE
with minor findings, but the release aims for a clean APPROVE in every domain.
If the overall pool is exhausted with an unresolved critical finding, publish
`v0.1.0-rc1` and document it in `KNOWN_ISSUES.md`.

| Domain | Used | Nominal remaining | Latest verdict | Notes |
|---|---:|---:|---|---|
| R-mech | 9 | 1 | APPROVE | P1 used three REVISE verdicts, then a finding-free APPROVE. P2 reviews #1–#4 found and remediated map/configuration and region-scaling majors. Fresh stateless P2 review #5 returned a clean APPROVE after independently verifying the component-first engine, exact spatial work, full suite, and staged evidence. |
| R-test | 4 | 6 | APPROVE | P1 used one REVISE and one APPROVE-with-minor. P2 review #1 found an external-link URL secret leak; review #2 independently returned a clean APPROVE after URI/metadata hardening and adversarial regressions. |
| R-repo | 3 | 7 | APPROVE | The first two P2 invocations remain charged as `REVISE`. Fresh review #3 returned `APPROVE` with one architecture-status minor: it still implied that R-test and all three P2 approvals were pending. This current-state update corrects that wording. Reserve at least 3 R-repo invocations for P9. |

Total verdicts used: 16 of 30. Pooled verdicts remaining: 14, including the
reserved minimum of 3 R-repo invocations for P9.

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
