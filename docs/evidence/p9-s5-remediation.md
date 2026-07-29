# P9 S5 and fixture-corpus remediation

Date: 2026-07-29.

The first combined P9 formal review approved mechanics but correctly blocked
the release because the then-current evidence showed that mandatory criterion
S5 failed and because F17 `unicode_names.xlsx` was absent. This remediation
keeps the frozen criterion and all historical P8 artifacts intact.

## Deterministic benchmark workload

`benchmarks/workloads.py` derives the evaluated workbooks from the canonical
F03, F07, F08, and F13 fixtures. It adds a synthetic 1,000-row, eight-column
`BenchmarkArchive` sheet. F03's Summary sheet receives the same history block
at row 1000 so B6's sheet-only baseline is measured against the same workload
shape. Both arms receive byte-identical workload files.

The builder performs surgical OOXML additions. It does not load and resave the
canonical workbook. The permanent regression compares every original ZIP
member byte-for-byte; only `[Content_Types].xml`, `xl/workbook.xml`,
`xl/_rels/workbook.xml.rels`, and F03's deliberately extended Summary sheet
may differ. A separate semantic comparison proves that every pre-existing F03
Summary cell, formula, and cached value remains unchanged. This matters because
an initial openpyxl-based implementation
erased cached error results and caused both Excel LSP B3 runs in an aborted
matrix to fail. That invalid partial artifact was deleted, the causal defect
was reproduced, and the complete matrix was restarted from zero.

## S5 result

The deterministic scripted replay uses pinned `tiktoken` `o200k_base` over
tool-result payload plus minimal call labels:

| Task | Excel LSP | Naive dump | Naive / Excel LSP |
|---|---:|---:|---:|
| B1 | 725 | 56,600 | 78.07× |
| B2 | 206 | 26,413 | 128.22× |
| B3 | 929 | 26,217 | 28.22× |
| B4 | 764 | 26,278 | 34.40× |
| B5 | 608 | 56,600 | 93.09× |
| B6 | 178 | 30,181 | 169.56× |
| **Total** | **3,410** | **222,289** | **65.19×** |

Every task and the aggregate exceed the frozen 10× threshold.

The fresh isolated headless Codex matrix in
`benchmarks/results/llm-eval-s5-remediation.jsonl` retains both repetitions:

- Excel LSP: 12/12 exact, 100.0%;
- naive dump: 8/12 exact, 66.7%;
- mean full CLI input-plus-output usage: 77,310.5 versus 64,909.8 tokens.

The full CLI value includes fixed agent instructions, tool schemas, and model
reasoning. It remains visible as a secondary measurement and is not described
as a reduction. S5's conjunction passes because the defined deterministic
workbook-payload metric exceeds 10× and LLM accuracy is better.

## Headless-run guard accounting

The 80-run guard counts every Codex invocation, including rows that were later
deleted and the invocation interrupted with Ctrl+C:

| Source | Retained | Completed | Interrupted | Counted runs |
|---|---:|---:|---:|---:|
| Historical P8 mixed preflight | yes | 24 | 0 | 24 |
| Historical P8 corrected baseline rerun | yes | 12 | 0 | 12 |
| Invalid openpyxl-workload P9 prefix | no | 12 | 1 | 13 |
| Fresh surgical-workload P9 matrix | yes | 24 | 0 | 24 |
| **Cumulative** |  | **72** | **1** | **73 / 80** |

The discarded prefix completed B1–B3 across both arms and repetitions. The
runner then started B4 Excel-LSP repetition 1, where the task transcript records
the KeyboardInterrupt. This fixes the earlier retained-row-only count of 60.
`benchmarks/analyze_results.py` derives all retained counts from their raw files
and adds the explicit 12-completed/one-interrupted deleted prefix before writing
the structured `run_guard` object in `results/audit-cost.json`. Seven runs
remain; no further model run was needed for this accounting correction.

## F17 corpus completion

`tests/fixtures/generate.py` now emits F17 `unicode_names.xlsx` with:

- worksheet names `Résumé d'été`, `東京`, and `O'Brien résumé`;
- correctly doubled apostrophe quoting in cross-sheet formulas;
- Unicode values and a Unicode defined name `TauxÉté`;
- deterministic formula caches.

Dedicated shape and semantic assertions cover raw OOXML and the production
parser. The openpyxl oracle independently observes the Unicode values,
formulas, and caches. The corpus has 21 numbered cases represented by 22 files
because F09a and F09b are distinct workbooks.

## Evidence commands

```console
uv run python tests/fixtures/generate.py
uv run python benchmarks/run_scripted.py
uv run python benchmarks/run_llm_eval.py --output benchmarks/results/llm-eval-s5-remediation.jsonl
uv run python benchmarks/analyze_results.py
uv run python benchmarks/plot.py
uv run pytest tests/unit/test_benchmark_checkers.py tests/unit/test_fixture_generation.py tests/oracle/test_oracle.py -q
```

The final broad test, coverage, formatting, type, packaging, CI, and clean
install results are recorded in `p9-release.md` after this candidate is fully
verified.
