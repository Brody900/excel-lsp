# Evidence

This directory is the audit trail for Excel LSP's implementation, tests, live
compatibility, benchmarks, installation, and release claims. Do not record
credentials, private workbook data, signed URLs, environment-variable values,
or other secrets.

## Verified evidence

- [Phase 0 reconnaissance](p0-recon.md) records the development environment,
  Excel/VBA probe, dependency pins, SQLite R*Tree support, package-name check,
  Codex headless probe, and scaffold verification.
- [Phase 1 parser and index foundation](p1-foundation.md) records parser,
  lifecycle, canonical-export, oracle, concurrency, coverage, build, and review
  evidence.
- [README claims-to-artifacts matrix](readme-claims-to-artifacts.md) maps every
  public claim or required README section to its producing phase and exact
  proof. P9 release claims remain `Candidate P9` until the formal gate closes.

## Completed phase evidence

- [Phase 2 regions, symbols, and workbook map](p2-regions-map.md) records the
  implemented P2 contracts, fixtures, invariants, and measured F03/F20 map
  budgets. Its R-mech, R-test, and user-authorized early R-repo gates have all
  approved.
- [Phase 3 formula references and blocks](p3-formulas-blocks.md) records the
  implemented reference, R1C1, block, fixture, lifecycle, privacy, and
  performance evidence. Fresh R-mech and R-test reviewers both cleanly approved
  the final frozen fingerprint.
- [P3 desktop-Excel declaration oracle](p3-excel-declaration-oracle.csv)
  preserves the exact 49 worksheet-entry probes behind the raw LET/LAMBDA name
  grammar decision (26 calculated, 23 rejected on Excel 16.0 build 19530).
- [Phase 4 dependency graph and traces](p4-graph.md) records the implemented
  graph, trace, path, circular, fixture, golden, and bounded-work evidence.
  Its combined formal reviewer cleanly approved both R-mech and R-test on one
  exact frozen fingerprint.
- [Phase 5 diagnostics](p5-diagnostics.md) records the complete catalog,
  error/link/volatile persistence, filtering, four fixtures, exact golden,
  oracle agreement, lifecycle invalidation, and verified combined formal gate.
- [Phase 6 surgical editor](p6-editor.md) records the verified OOXML writer,
  direct index patch, staleness lifecycle, R1C1 column fill, complete F16/F21
  part manifests, property proof, and desktop-Excel smoke. Its combined formal
  reviewer approved both gates on the exact final frozen fingerprint.
- [Phase 7 MCP server and CLI](p7-mcp-cli.md) records the 14-tool stdio
  service, exact generated schemas, shared CLI service, progress, caps, cursor
  invalidation, regex timeout, path confinement, and verification. Its combined
  formal reviewer approved both gates on the exact final frozen fingerprint.

## Verified Phase 8 evidence

- [Phase 8 live desktop-Excel evidence](live-excel/index.md) records all seven
  required protocol steps, numbered screenshots, COM assertions, exact product
  responses, the VBA run, write refusal, chart/image preservation, three trace
  cross-checks, and the real-capture demo GIF.
- [Phase 8 benchmarks](p8-benchmarks.md) records the exact-graded six-task
  harness, both headless-Codex repetitions, raw measurements, S1 pass, honest
  S5 failure, optional-arm skip, and generated charts. Its combined reviewer
  approved both gates on the exact final frozen fingerprint.
- [Success criteria](success-criteria.md) cross-checks the frozen S1–S7 wording,
  retains the historical P8 failure, and records the candidate P9 remediation.

## Candidate Phase 9 evidence

- [Fresh-install evidence](fresh-install.md) records package build, wheel and
  `uvx` execution, MCP initialization, Codex registration, runtime network
  denial, and the public-git fallback.
- [Codex MCP help](codex-mcp-help.txt) records the release-time syntax and
  isolated registration result.
- [Comparison sources](comparison-sources.md) pins and dates every upstream
  observation used by the README comparison.
- [Registry submission packet](../registry-submissions.md) contains exact
  account-scoped follow-up copy and current submission constraints.
- [Phase 9 release evidence](p9-release.md) summarizes the packaging, public
  repository, verification, and remaining formal/release gates.
- [P9 S5 and fixture remediation](p9-s5-remediation.md) records the disclosed
  workload, fresh exact-answer matrix, passing mandatory criterion, and F17
  corpus completion.
