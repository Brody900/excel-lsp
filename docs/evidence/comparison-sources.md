# Comparison sources

Accessed **2026-07-29**. This comparison records what each project documents at
one pinned Git revision; it does not claim that an undocumented capability is
impossible or that the projects have identical goals.

## Pinned projects and scope

- **Excel LSP:** this repository's v0.1.0 candidate and the phase evidence
  linked below.
- **haris-musa/excel-mcp-server:**
  [`f51340e`](https://github.com/haris-musa/excel-mcp-server/tree/f51340ecd5778952405044b203d3a2d4c8a46833),
  with the pinned
  [README](https://github.com/haris-musa/excel-mcp-server/blob/f51340ecd5778952405044b203d3a2d4c8a46833/README.md).
  Its documented focus is broad workbook creation and manipulation, including
  formatting, charts, pivot tables, tables, and three transports.
- **jwadow/mcp-excel:**
  [`eb088c5`](https://github.com/jwadow/mcp-excel/tree/eb088c5edd5335c67ffc14e521be607a46d49b2a),
  with the pinned
  [README](https://github.com/jwadow/mcp-excel/blob/eb088c5edd5335c67ffc14e521be607a46d49b2a/README.md).
  Its documented focus is read-only atomic analytics with smart caching and
  context-protection limits; write operations are on its roadmap.
- **Naive dump:** the intentionally small
  [`benchmarks/baseline_server.py`](../../benchmarks/baseline_server.py)
  reference arm. It returns full worksheet CSV and has no semantic index.

The two upstream projects solve valuable but different problems. Excel LSP
prioritizes formula structure, navigation, diagnostics, and conservative
mutation fidelity; this table should not be read as an overall ranking.

## Persistent semantic index

Excel LSP persists a derived SQLite semantic index, verified in
[P1](p1-foundation.md#delivered-contracts). The pinned haris-musa README does
not document a persistent semantic index. The pinned jwadow README documents
in-process smart caching but not a persistent symbol/formula index. The naive
arm has neither.

## Formula dependency graph

Excel LSP's stored, bidirectional dependency graph and bounded navigation are
verified in [P4](p4-graph.md#delivered-contracts). Neither pinned upstream
README documents formula-precedent/dependent graph navigation. The naive arm
only renders formulas and cached values into CSV.

## Incremental reindex

Excel LSP hashes package parts and can reindex only affected worksheets; the
contract is verified in [P1](p1-foundation.md#invariant-evidence) and measured
in [`index-timing.csv`](../../benchmarks/results/index-timing.csv). Neither
pinned upstream README documents part-hash-driven incremental reindexing. The
naive arm reopens and renders workbook content on each request.

## Formula diagnostics

Excel LSP exposes typed diagnostics for cached errors, broken links,
inconsistent formulas, dynamic references, volatility, staleness, and related
conditions, verified in [P5](p5-diagnostics.md#diagnostic-matrix). Neither
pinned upstream README documents an equivalent formula-diagnostic catalog. The
haris-musa README does advertise range/formula/data-integrity validation, which
is a different and useful capability.

## Edit support and fidelity

Excel LSP has two deliberately narrow write tools and byte-identity evidence
for every untouched OOXML part in F16 and F21. The haris-musa project documents
broad write support, including workbook, formatting, chart, pivot, and table
operations; its pinned README does not claim untouched-part byte identity. The
jwadow project documents XLS/XLSX as read-only and lists writes on its roadmap.
The naive arm is read-only.

Evidence: [P6 part preservation](p6-editor.md#part-preservation),
[`part-diff-f16.json`](part-diff-f16.json),
[`part-diff-f21.json`](part-diff-f21.json), and the
[live assertions](live-excel/index.md#machine-readable-assertions).

## Token discipline

Excel LSP enforces per-response value and serialized-character caps, but the P8
suite did **not** show token reduction: Excel LSP used more tokens than the
naive arm. That unfavorable result remains explicit in the
[S5 calculation](success-criteria.md#s5). The pinned jwadow README documents
context-protection limits and bounded previews; the pinned haris-musa README
does not state comparable serialized response caps. No cross-project token
benchmark was run, so no competitor token-efficiency ranking is claimed.
