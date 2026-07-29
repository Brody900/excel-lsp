# Benchmarks

Excel LSP's benchmark suite is reproducible, exact-graded, and intentionally
keeps secondary and unfavorable results visible. It compares:

- **excel-lsp** — the persistent semantic index and bounded MCP tools;
- **naive-dump** — the two-tool baseline in `baseline_server.py`, which returns
  a complete workbook or sheet as CSV text.

Each benchmark workbook is built by `workloads.py` from its canonical F03,
F07, F08, or F13 fixture. The builder surgically adds a disclosed 1,000-row,
eight-column `BenchmarkArchive` region; F03's Summary sheet receives the same
history block so the sheet-only B6 baseline is measured fairly. Tests prove
that every canonical OOXML member remains byte-identical except the package
declarations needed for the added sheet and F03's deliberately extended
Summary XML. Both arms receive the exact same workbook.

The optional `haris-musa/excel-mcp-server` arm was probed with
`uvx excel-mcp-server stdio`. It installed and listed 25 tools in 2.6 seconds,
but its own stdio entry point raised `ValueError: I/O operation on closed file`
on normal client shutdown. HANDOFF §9.1 permits this arm only when it installs
and runs cleanly within 15 minutes, so it is excluded rather than represented
as a DNF or silently treated as successful.

## Task and grading contract

The six tasks in [`tasks/`](tasks/) cover lineage, formula-pattern audit,
cached errors, schema inference, impact analysis, and a bounded QA lookup.
Every prompt requires a final `ANSWER: <json>` line. `check.py` parses only
that final line. It exact-compares scalar/object tasks and compares arrays as
duplicate-free sets for the four tasks whose answers are inherently
order-insensitive. The public Markdown contracts are rendered from the same
`TaskSpec` objects used to create the evaluated prompts. No prose or fuzzy
grader affects the score.

## Measurement modes

`uv run excel-lsp bench` (or `uv run python benchmarks/run_scripted.py`) runs
the deterministic scripted replay. It measures only tool-result payload plus
minimal call labels using pinned `tiktoken` and `o200k_base`.

`uv run python benchmarks/run_llm_eval.py` runs isolated Codex CLI agents with
model `gpt-5.6-sol`, reasoning effort `high`, read-only sandboxing, approval
policy `never`, project rules disabled, and only the task-specific MCP tools
enabled. Each task/arm cell has two repetitions. The runner records every raw
JSON event, final transcript, exact grade, tool calls, wall time, CLI token
usage, protocol violations, return code, and any reported dollar cost. It is
resumable and stops before 80 runs or $15 of reported spend.

Codex CLI 0.144.5 emitted input/output usage but no dollar-cost field under the
available ChatGPT authentication, so the committed cost callout reports tokens
and time and marks dollars unavailable. It does not invent a conversion.

## LLM repetitions and agreement

`results/llm-eval.jsonl` contains the fresh S5-remediation 6 × 2 × 2 matrix.
Its untouched raw source is `results/llm-eval-s5-remediation.jsonl`. Both
repetitions remain separate in `results/accuracy.csv`; `results/accuracy.md`
renders the exact-answer table and marks agreement by parsed JSON value.

Historical P8 inputs remain committed for audit. The initial mixed preflight
exposed a baseline-server annotation defect:
read-only tools without `readOnlyHint=true` were canceled by approval policy
`never`. Those invalid baseline rows remain in
`results/llm-eval-preflight-mixed.jsonl`; the corrected baseline reruns remain
in `results/llm-eval-baseline-rerun.jsonl`. They are historical audit inputs
only. `analyze_results.py` now selects both arms from the fresh
`llm-eval-s5-remediation.jsonl`, verifies the exact matrix, and emits the
consolidated artifacts. The original model transcripts remain unchanged. When
the corrected order-insensitive checker changes a stored grade, the
consolidated row retains `source_grade` and marks
`regraded_without_model_rerun=true` so the correction is auditable.

The run guard counts every invocation, not only retained rows: 24 historical
preflight runs, 12 historical baseline reruns, 12 completed plus one
interrupted run from the deleted invalid-workload matrix, and 24 retained P9
runs. The exact cumulative total is 73/80; the structured breakdown is emitted
in `results/audit-cost.json`.

Before publication, the runner replaces the local checkout path with
`<WORKSPACE>` throughout transcripts, stderr, and nested raw events. The same
sanitizer was applied to the pre-existing raw files; it does not alter answers,
usage, timing, tool calls, or model output other than that path redaction.

## Reproduction

From a clean checkout with `uv` available:

```powershell
uv sync --locked
uv run python tests/fixtures/generate.py
uv run excel-lsp bench
uv run python benchmarks/run_llm_eval.py --output <new-jsonl-path>
uv run python benchmarks/analyze_results.py
uv run python benchmarks/run_index_timing.py
uv run python benchmarks/plot.py
uv run pytest tests/unit/test_benchmark_checkers.py tests/unit/test_benchmark_llm_runner.py tests/unit/test_benchmark_analysis.py
```

Do not overwrite the committed LLM evidence when reproducing; choose a new
output path. Headless runs consume account capacity even when no dollar field
is reported.

## Results

The P9 remediation preserves S1 and now satisfies both halves of S5:

- 50,000 × 10 cold median: **9.440 s**; one-Control-sheet incremental median:
  **0.066 s**.
- Headless exact accuracy: **12/12 (100.0%)** for Excel LSP and **8/12 (66.7%)**
  for naive dump.
- Scripted tool-result payload totals: **3,410** vs **222,289** tokens; naive
  dump used **65.2×** as many workbook-payload tokens. Every task individually
  exceeds the frozen 10× threshold.
- Mean full Codex usage: **77,310.5** vs **64,909.8** tokens. This secondary
  measurement includes fixed agent context, MCP schemas, and reasoning, so it
  is reported separately rather than substituted for the workbook-payload
  metric named by S5.

See [`docs/evidence/p8-benchmarks.md`](../docs/evidence/p8-benchmarks.md) and
[`docs/evidence/success-criteria.md`](../docs/evidence/success-criteria.md) for
the commands, calculations, limitations, and criterion status.
