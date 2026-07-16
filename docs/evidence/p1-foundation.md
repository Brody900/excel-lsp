# Phase 1 parser, index, and lifecycle evidence

Recorded on 2026-07-15 (America/Los_Angeles). This phase implements M1 and M8:
the streaming OOXML package reader, typed normalized cell stream, selected-part
hashing, SQLite store/spatial fallback, and freshness lifecycle.

## Delivered contracts

- Namespace-tolerant workbook, relationship, content-type, shared-string,
  style, worksheet, external-link, ListObject, defined-name, `calcPr`, VBA,
  array, shared-formula, and What-If Data Table parsing.
- Callback-driven `lxml.iterparse` worksheet reads with element clearing and no
  production openpyxl workbook load/save path.
- One public JSON-scalar normalization boundary and int64-safe SQLite numeric
  persistence for Excel's finite floating-point domain.
- Whole-file and selected-part SHA-256 hashes, including worksheet relationship
  and ListObject dependencies.
- WAL-backed SQLite schema, R*Tree/interval abstraction, natural-key canonical
  exports, generation counters, exact sidecar placement, full/per-sheet/no-op
  indexing, and per-call freshness entry point.
- Structured unsupported, missing, encrypted, locked, and corrupt errors,
  including locked-open and torn-save retry-once paths.

## Invariant evidence

| Invariant | Evidence |
|---|---|
| I1 untouched reindex is a no-op | Fake-parser lifecycle and real F01 integration tests assert unchanged generation, no parser calls on the stat fast path, and no reindexed sheets. Byte-identical mtime-only and invalid-stat refreshes take a hash no-op, repair bookkeeping without invalidating generation-bound cursors, and retry if the source stat changes again during hashing. |
| I2 full and incremental canonical exports agree | A changed-sheet lifecycle test builds independent incremental and full databases and compares all natural-key exports. Incoming edges targeting a refreshed sheet are preserved. Cross-database tests vary physical sheet and edge ids and prove equal spatial exports for both R*Tree and interval backends while retaining each edge-to-rectangle association. |
| I3 source is never mutated | Parser and lifecycle tests compare source bytes/hash/stat before and after parse, index, and no-op refresh. |

The parser derives actual bounds from streamed cells rather than trusting
`<dimension>`. Tests cover false/missing dimensions, rich and inline strings,
all specified scalar types, 1900/1904 dates including serial 60, shared and
array formulas, What-If Data Tables, validations, merges, tables, chart sheets,
external links, VBA, nonstandard namespaces, hashes, malformed packages, and
source immutability.

## Deterministic fixture and oracle evidence

Consecutive generator runs produced byte-identical workbooks:

```text
basic_single_table.xlsx 8d57d9143edf78a66be6c33bcede3bcc7fba8ed1ac2d816391a4139a28a41270
formula_blocks.xlsx 50015028edc75a4bab5cd13af9b4576f520d8c1f4cf0e3b223bab54c3476c871
```

Both emitted fixtures have exact production-parser versus pinned openpyxl
dual-load equality for `(sheet, ref, cached value, formula)`. F07 independently
asserts all 20 injected cached values, both genuine shared groups, the planted
tamper, ListObject, and merge.

The informational `openpyxl==3.1.5` probe observed:

- shared followers C3, C11, C14, and C21 expand to their translated formulas;
- `ReadOnlyWorksheet.tables` raises `AttributeError`;
- `ReadOnlyWorksheet.merged_cells` raises `AttributeError`;
- a normal-mode control sees `FormulaBlocksTable` and `E1:F1`.

There are no active cell-stream oracle skips. Reproduce with either:

```text
uv run python tests/oracle/probe_openpyxl.py
uv run python -m tests.oracle.probe_openpyxl
```

## Concurrency and incremental probes

- Five in-suite rounds of eight simultaneous cold opens completed with one
  valid schema and no exceptions.
- Six simultaneous opens of an obsolete schema performed exactly one migration
  and all observed generation 8.
- The mechanics reviewer's 30-database, eight-thread cold-open probe was rerun
  after the fix: `failed_runs=0`, with no captured failures.
- Table-name-only and table-ref-only mutations leave worksheet XML byte-identical,
  change only the referenced table metadata among the owning sheet's indexed
  dependencies, and reindex exactly that sheet.

## Fresh orchestrator verification

| Check | Result |
|---|---|
| `uv lock --check` | passed; 70 packages resolved |
| `uv run python tests/fixtures/generate.py` | passed; F01 and F07 regenerated |
| `uv run ruff check .` | passed |
| `uv run ruff format --check .` | passed; 33 files formatted |
| `uv run pyright` | passed; 0 errors, warnings, or information messages |
| `uv run pytest --cov=excel_lsp.core --cov-report=term-missing --cov-fail-under=85` | 160 passed; 85.75% branch coverage |
| `uv build` | built `excel_lsp-0.1.0` sdist and wheel |

## Review accounting

Verdict-bearing invocations:

- R-mech #1: REVISE; four major findings covering Data Table semantics,
  table-part invalidation, incoming edges, and concurrent schema creation.
- R-mech #2: REVISE; three major findings covering spatial edge association,
  content-type invalidation, and duplicate-row corruption shaping.
- R-mech #3: REVISE; two major findings covering mtime-only generation bumps
  and physical sheet ids in spatial canonical exports.
- R-mech #4: clean APPROVE; no findings after targeted adversarial inspection
  and 153 focused tests.
- R-test #1: REVISE; three major findings covering `calcPr`, locked-open retry,
  and parse-time torn-save retry evidence.
- R-test #2: APPROVE with one documentation-accounting minor, corrected before
  phase close; its fresh run observed 150 passing tests and 86.06% coverage.

Two read-only reviewer processes ended before returning the strict review
protocol output: one was stopped after an unbounded delegated probe and one was
lost during an automatic thread restart. They produced no verdict or gate
artifact and are recorded in the agent log as operational interruptions rather
than review-ledger invocations.

Future formula classification, regions, graph construction, diagnostics,
editor, server caps/cursors, full fixture corpus, live Excel evidence, and
benchmarks remain assigned to their ordered P2-P9 gates.
