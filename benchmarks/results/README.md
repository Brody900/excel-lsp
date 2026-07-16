# Benchmark results

Commit raw benchmark measurements here together with enough environment metadata
to reproduce them.

## Current artifact

- `map-budgets.json` — P2's deterministic normalized workbook-map character and
  `o200k_base` token counts for F03 and F20. The exact golden-map test must
  reproduce every recorded value.

## Required P8 artifacts

- `environment.json` — git revision, OS, Python, dependency versions, CPU,
  Codex CLI/model details, verified command-line flags, token encoding, and run
  timestamps.
- `scripted.csv` — one checked row per deterministic task and benchmark arm,
  including payload tokens, tool-call count, wall time, and the final exact
  `ANSWER:` result.
- `llm-eval.jsonl` — both isolated headless-Codex repetitions for every attempted
  task/arm cell, retaining raw usage and cost fields, exact-answer output,
  status, and any explicit DNF reason.
- `accuracy.csv` — checker result for every scripted and LLM-eval run plus the
  two-repetition agreement field; prose or fuzzy grading is not accepted.
- `index-timing.csv` — cold and one-sheet incremental index timings for the
  1,000-, 10,000-, and 50,000-row fixtures with environment identifiers.
- `audit-cost.json` — the computed cost-of-one-audit value and the exact
  `llm-eval.jsonl` rows from which it was derived.

## Acceptance gates

P8 may treat these files as benchmark evidence only when:

1. `benchmarks/check.py` accepts every completed transcript's final `ANSWER:`
   line and the committed accuracy rows agree with the checker output.
2. `excel-lsp bench` runs the documented reproducible harness, proven by
   `tests/unit/test_cli.py::test_bench_command_runs_reproducible_harness` and
   `docs/evidence/p8-benchmarks.md#excel-lsp-bench`.
3. Both headless-Codex repetitions are reported individually, disagreements and
   DNFs remain visible, and the cost guard is documented.
4. `benchmarks/plot.py` regenerates every PNG/SVG chart from these committed raw
   files without hand-entered values.
5. `docs/evidence/success-criteria.md` records the S1 and S5 calculations using
   the frozen criterion: “Benchmarks show ≥ 10× token reduction vs. the
   naive-dump baseline on the defined task suite, with equal-or-better task
   accuracy in LLM evals.”

Until those gates pass, the P8 filenames above are a release contract, not
measurements or performance claims.
