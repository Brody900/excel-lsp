# Success criteria

This ledger preserves the frozen HANDOFF §1 criteria. Historical failures and
secondary measurements remain visible rather than being renamed or hidden.

| Criterion | Current status | Evidence |
|---|---|---|
| S1 | **Pass (verified P8)** | `benchmarks/results/index-timing.csv`; 50k × 10 median 9.439544 s cold and 0.065912 s after a one-sheet change |
| S2 | **Pass (verified P2)** | `benchmarks/results/map-budgets.json`; `docs/evidence/p2-regions-map.md#map-budgets` |
| S3 | **Pass (verified P4)** | `tests/golden/p4-graph-semantics.json`; `docs/evidence/p4-graph.md#fixture-and-golden-evidence` |
| S4 | **Pass (verified P8)** | `docs/evidence/live-excel/index.md`; P6 byte-level F16/F21 preservation plus desktop-Excel authoring, recalc, VBA run, write refusal, and chart/image evidence |
| S5 | **Pass (candidate P9)** | Deterministic payload 3,410 vs 222,289 tokens (65.2× reduction; every task ≥10×); headless accuracy 100.0% vs 66.7% |
| S6 | **Pass (verified P7)** | `docs/evidence/p7-mcp-cli.md#response-caps`; real stdio and focused boundary tests |
| S7 | **Pass (candidate P9)** | Clean wheel and local-wheel `uvx` probes; public-git `uvx --from` fallback; Codex registration evidence |

## S1

The final timing series has three repetitions at each size. At 50,000 rows and
10 `Perf` columns, cold samples are 9.213166, 9.439544, and 9.445920 seconds;
their median is **9.439544 seconds**, below the strict 10-second gate.
Incremental samples are 0.065652, 0.065912, and 0.069194 seconds; their median
is **0.065912 seconds**, below the strict 1-second gate.

Cold timing indexes the complete workbook (`Perf` plus a tiny `Control` sheet).
Incremental timing changes only `Control!A2` by replacing
`xl/worksheets/sheet2.xml` in a benchmark copy, excludes mutation time, and
requires the lifecycle result to report exactly one reindexed sheet.

## S5

The frozen criterion requires both ≥10× token reduction against naive dump and
equal-or-better LLM accuracy.

Accuracy passes: Excel LSP produced 12 exact answers out of 12 (**100.0%**),
while naive dump produced 8 out of 12 (**66.7%**).

Token reduction passes on the frozen deterministic tool-result payload metric:

- scripted totals are **3,410** for Excel LSP and **222,289** for naive dump;
  naive dump therefore carries **65.187×** as many workbook-payload tokens;
- every B1–B6 task individually exceeds 10× (the smallest ratio is 28.22×);
- both arms receive the same canonical task workbook plus a deterministic,
  disclosed 1,000-row archive workload; a permanent test proves every original
  OOXML member remains byte-identical except the declarations for the added
  sheet and F03's deliberately extended Summary XML.

The secondary full-agent measurement is also published: mean input-plus-output
CLI usage is 77,310.5 tokens for Excel LSP and 64,909.8 for naive dump. It
includes fixed Codex context, tool schemas, and reasoning and is therefore not
substituted for the tool-result payload criterion. With ≥10× payload reduction
and equal-or-better exact LLM accuracy both true, **S5 passes**. The original P8
failure remains in `docs/evidence/p8-benchmarks.md` and the historical raw runs.

## S4

The P6 part-diff and property evidence proves untouched OOXML members remain
byte-identical. The P8 live pass adds the required desktop proof: three
Excel-authored workbooks were indexed; exact diagnostics and lineage were
cross-validated; L1 and F16 opened without repair and recalculated after both
write tools; `Stamp` ran and wrote 42 to Z1; an open L1 refused a product write;
and F21 retained its rendered chart and embedded image. All numbered captures
and exact JSON responses are indexed in `docs/evidence/live-excel/index.md`.

## S7

The wheel installed into a new Python 3.12 environment, mapped F03, initialized
over stdio, listed all 14 tools, and opened the workbook. The same checks passed
through `uvx --from` the locally built wheel. A separate public-git fallback
built the newest public commit directly from GitHub and ran the CLI and MCP
surface. Exact commands and environment metadata are in
[`fresh-install.md`](fresh-install.md). This satisfies the frozen fallback
clause without claiming that the still-unpublished PyPI artifact exists.
