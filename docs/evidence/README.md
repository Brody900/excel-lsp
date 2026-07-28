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
- [README claims-to-artifacts plan](readme-claims-to-artifacts.md) maps every
  public claim or required README section to its producing phase and exact
  proof. At P2, most later-phase rows are intentionally marked `Planned`.

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

## Active phase evidence

Phase 5 diagnostics is now the active implementation phase; its evidence report
will be created with that phase's first verified candidate results.

## Planned evidence paths

These paths are contracts, not links to completed work:

- `p5-diagnostics.md`: complete diagnostics matrix.
- `p6-editor.md`: surgical editing and staleness behavior.
- `part-diff-f16.json` and `part-diff-f21.json`: untouched-part fidelity.
- `p7-mcp-cli.md`: 14-tool conformance, annotations, instructions, caps, and CLI.
- `live-excel/index.md`: numbered live protocol, screenshots, and demo capture.
- `success-criteria.md`: final S1-S7 evidence cross-check.
- `fresh-install.md`: clean package, CLI, Codex MCP, and fallback verification.
- `codex-mcp-help.txt`: sanitized release-time Codex syntax capture.
- `comparison-sources.md`: pinned, dated sources for every comparison cell.

Each future evidence document must identify the exact commit, environment,
commands, results, and underlying raw artifact. A passing test source without a
fresh run, or a chart without raw rows and its generation script, is not enough
to mark a claim verified.
