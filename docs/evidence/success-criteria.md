# Success criteria

This ledger preserves the frozen HANDOFF §1 criteria. A failure is not renamed,
weakened, or hidden to make the release look complete.

| Criterion | Current status | Evidence |
|---|---|---|
| S1 | **Pass (verified P8)** | `benchmarks/results/index-timing.csv`; 50k × 10 median 9.439544 s cold and 0.065912 s after a one-sheet change |
| S2 | **Pass (verified P2)** | `benchmarks/results/map-budgets.json`; `docs/evidence/p2-regions-map.md#map-budgets` |
| S3 | **Pass (verified P4)** | `tests/golden/p4-graph-semantics.json`; `docs/evidence/p4-graph.md#fixture-and-golden-evidence` |
| S4 | **Pass (verified P8)** | `docs/evidence/live-excel/index.md`; P6 byte-level F16/F21 preservation plus desktop-Excel authoring, recalc, VBA run, write refusal, and chart/image evidence |
| S5 | **Fail (verified P8)** | Excel LSP accuracy 100.0% vs 75.0%, but scripted payload 3,375 vs 2,127 tokens and full Codex mean 77,927.8 vs 41,432.8 |
| S6 | **Pass (verified P7)** | `docs/evidence/p7-mcp-cli.md#response-caps`; real stdio and focused boundary tests |
| S7 | **Pending P9** | Clean `uvx` or documented git-fallback evidence not yet produced |

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
while naive dump produced 9 out of 12 (**75.0%**).

Token reduction fails:

- scripted totals: 3,375 / 2,127 = **1.5867× baseline usage**; equivalently,
  the baseline is 0.6302× Excel LSP, not Excel LSP being ≤0.1× baseline;
- headless means: 77,927.833 / 41,432.833 = **1.8808× baseline usage**;
  equivalently, the baseline is 0.5317× Excel LSP.

Because the required conjunction is false, **S5 fails**. The small deterministic
fixtures make a complete CSV dump unusually compact, while semantic responses
and MCP schemas have fixed structure; this explains the result but does not
change it. The project must not advertise a token reduction from this suite.

## S4

The P6 part-diff and property evidence proves untouched OOXML members remain
byte-identical. The P8 live pass adds the required desktop proof: three
Excel-authored workbooks were indexed; exact diagnostics and lineage were
cross-validated; L1 and F16 opened without repair and recalculated after both
write tools; `Stamp` ran and wrote 42 to Z1; an open L1 refused a product write;
and F21 retained its rendered chart and embedded image. All numbered captures
and exact JSON responses are indexed in `docs/evidence/live-excel/index.md`.

## Remaining criteria

S7 belongs to P9 and requires a clean environment. This document will be
refreshed at each phase boundary without retroactively altering the measured
S1/S5 rows.
