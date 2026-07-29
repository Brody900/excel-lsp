# Benchmark results

These are committed measurements, not hand-entered marketing numbers. The
generation scripts live one directory above; charts are generated from the
CSV/JSON artifacts and committed under `docs/assets/` as PNG and SVG.

| Artifact | Purpose |
|---|---|
| `environment.json` | OS, CPU, Python, Codex CLI/model/reasoning, tokenizer, timestamp, and timing caveats |
| `scripted.csv` | Twelve deterministic task/arm replays with payload tokens, calls, time, exact answer, and checker result |
| `llm-eval.jsonl` | Consolidated, exact 6 × 2 × 2 headless-Codex matrix with raw events and transcripts |
| `llm-eval-s5-remediation.jsonl` | Fresh raw 6 × 2 × 2 matrix on the disclosed deterministic archive workload |
| `accuracy.csv` | Flat per-repetition usage, status, correctness, agreement, call count, and reported cost |
| `accuracy.md` | Human-readable exact-answer accuracy table |
| `audit-cost.json` | B2 Excel-LSP mean token/time callout and explicit unavailable-dollar note |
| `index-timing.csv` | Three cold and incremental repetitions for 1k, 10k, and 50k rows |
| `index-timing-pre-optimization.csv` | Retained before/after engineering evidence, including the disclosed desktop-suspension outlier |
| `llm-eval-preflight-mixed.jsonl` | Historical P8 mixed run retained for audit |
| `llm-eval-baseline-rerun.jsonl` | Historical P8 corrected naive-baseline reruns retained for audit |
| `map-budgets.json` | Earlier deterministic F03/F20 workbook-map character and token budgets |

`benchmarks/analyze_results.py` rejects missing/duplicate matrix cells and
auditably preserves any source grade changed by the corrected checker.
`benchmarks/check.py` enforces the final-line `ANSWER:` contract plus each
task's exact or duplicate-free set semantics. Raw transcripts and usage remain
unchanged; local checkout paths are replaced with `<WORKSPACE>` before public
storage. Run `excel-lsp bench` to execute the documented reproducible harness.
The frozen S5 criterion begins “Benchmarks show ≥ 10× token reduction vs. the
naive-dump baseline.” The primary deterministic payload measure is 3,410
tokens for Excel LSP versus 222,289 for naive dump, a 65.2× reduction.

## Timing interpretation

F06 has a 10-column `Perf` sheet plus a small `Control` sheet. Cold timing
indexes both sheets. Incremental timing surgically changes only
`xl/worksheets/sheet2.xml` (`Control!A2`) and requires exactly one reindexed
sheet. Mutation time is excluded. Filesystem cache state was not destructively
cleared and is disclosed in `environment.json`.

At 50,000 rows, the cold samples are 9.213166, 9.439544, and 9.445920 seconds;
the incremental samples are 0.065652, 0.065912, and 0.069194 seconds. Their
medians satisfy S1's strict 10-second and 1-second thresholds.

## Acceptance status

- Exact graders accept every scripted answer and were rerun over every
  consolidated LLM transcript.
- Both LLM repetitions and disagreements are visible.
- The run guard records all **73** headless invocations—retained, failed,
  deleted, and interrupted—in a structured breakdown, leaving 7 below its
  80-run ceiling.
- Five PNG/SVG chart pairs regenerate from committed artifacts.
- S1 passes.
- S5 passes: every deterministic task exceeds 10× payload reduction and the
  aggregate is 65.2×, while headless accuracy is 100.0% versus 66.7%.
- Full CLI usage remains separately visible and is not presented as a 10× win.
