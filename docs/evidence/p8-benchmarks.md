# Phase 8 benchmark evidence

Status: verified P8 evidence; the combined formal reviewer approved both gates
on staged tree `e299db5fed3d26db9b85b838f4476e1a86692efd`.

Environment: Windows 11, AMD64 Family 25 Model 117, Python 3.12.11,
Codex CLI 0.144.5, `gpt-5.6-sol`, reasoning effort `high`, and pinned
`tiktoken` `o200k_base`. The raw capture is
[`environment.json`](../../benchmarks/results/environment.json).

## Delivered harness

- `benchmarks/model.py` freezes B1–B6 fixtures, questions, JSON shapes, and
  expected values.
- `benchmarks/check.py` parses only the final `ANSWER:` line, exact-grades
  scalar/object tasks, and uses duplicate-free set comparison for inherently
  order-insensitive answer arrays.
- `benchmarks/baseline_server.py` implements the naive complete-workbook and
  complete-sheet CSV tools with explicit read-only MCP annotations.
- `benchmarks/run_scripted.py` runs deterministic reasonable call sequences and
  counts result payloads with pinned `o200k_base`.
- `benchmarks/run_llm_eval.py` runs isolated, task-restricted, read-only Codex
  CLI sessions with resumability, timeout handling, protocol auditing, and the
  80-run/$15 guard.
- `benchmarks/analyze_results.py` validates the exact 24-row matrix and
  auditably retains source grades when a corrected checker changes one before
  writing consolidated evidence.
- `benchmarks/run_index_timing.py` produces the three-repetition cold and
  one-sheet incremental series.
- `benchmarks/plot.py` regenerates five PNG/SVG pairs without hand-entered
  measurements, with a stable SVG hash salt so unchanged inputs reproduce
  byte-identical PNG and SVG assets.

## `excel-lsp bench`

`uv run excel-lsp bench` calls the same scripted runner used to create
`benchmarks/results/scripted.csv`. The permanent CLI test redirects the output
to a temporary path and requires a JSON summary of 12 rows and zero failures:

```text
tests/unit/test_cli.py::test_bench_command_runs_reproducible_harness
```

This establishes that the documented product command, not only an internal
script, runs the reproducible harness.

## Headless Codex isolation and raw evidence

The installed syntax was verified from `codex exec --help` and captured in
[`codex-exec-help.txt`](codex-exec-help.txt). Runs used:

- approval policy `never` and read-only sandbox;
- `--ephemeral`, `--ignore-rules`, and `--skip-git-repo-check`;
- zero project-document bytes;
- user MCP servers and unrelated capabilities disabled;
- only task-specific tools enabled on the selected benchmark server;
- a 300-second per-run timeout and JSON event output.

`--ignore-user-config` was tested but could not be used with the available
ChatGPT authentication because it produced HTTP 401. The runner instead keeps
authentication while explicitly disabling user MCPs and unrelated features.
No secrets, auth headers, or configuration values are stored. The runner
recursively replaces the local checkout path with `<WORKSPACE>` in transcripts,
stderr, and raw events; the same sanitizer was applied to the already-captured
raw files without changing answers, usage, timing, or tool calls.

The original naive server omitted `readOnlyHint`; with approval policy `never`,
Codex correctly canceled tools it could not identify as non-mutating. This was
an evaluation-server annotation defect, not a workbook-content or product
cybersecurity finding. The invalid baseline rows remain in
`llm-eval-preflight-mixed.jsonl`. After adding `readOnlyHint=true` and
`openWorldHint=false`, twelve clean baseline reruns were recorded. The
consolidator selects only Excel-LSP rows from the original run and only baseline
rows from the rerun, then regrades all 24 selected transcripts.

## Results

### Scripted payload

| Arm | Six-task total | Mean per task |
|---|---:|---:|
| Excel LSP | 3,375 | 562.5 |
| Naive dump | 2,127 | 354.5 |

Excel LSP used 1.5867× the baseline payload (58.7% more), so this mode does not
show a token reduction.

### Headless Codex

| Arm | Mean total CLI tokens | Exact answers | Accuracy |
|---|---:|---:|---:|
| Excel LSP | 77,927.8 | 12/12 | 100.0% |
| Naive dump | 41,432.8 | 9/12 | 75.0% |

Excel LSP used 1.8808× the baseline's full CLI tokens (88.1% more). Its accuracy
was higher by three exact answers. The first grading pass incorrectly treated
B5's inherently set-valued dependent ranges as an ordered array even though
the evaluated prompt never specified an order. Both Excel-LSP repetitions had
returned the exact correct set. The corrected checker makes B1, B2, B3, and B5
order-insensitive while rejecting duplicates; B4 and B6 remain exact. No
evaluated prompt or model transcript changed, so no model rerun was needed.
Every changed consolidated row retains its original `source_grade` and
`regraded_without_model_rerun=true`. The naive arm still disagreed across B5
repetitions. Every per-run status is visible in
[`accuracy.md`](../../benchmarks/results/accuracy.md).

### Index timing

| Perf rows × columns | Cold median | Incremental median |
|---|---:|---:|
| 1,000 × 10 | 0.186840 s | 0.022772 s |
| 10,000 × 10 | 1.699052 s | 0.030830 s |
| 50,000 × 10 | 9.439544 s | 0.065912 s |

The complete cold workbook also contains a two-cell `Control` sheet. The
incremental mutation changes only that sheet and exactly one sheet is
reindexed. The pre-optimization measurements remain committed; one incremental
sample was prolonged by desktop disconnect/suspension and is explicitly not
used in the final series.

## Optional arm C

The bounded probe `uvx excel-mcp-server stdio` installed and exposed 25 tools,
including `read_data_from_excel` and `get_workbook_metadata`, in 2.6 seconds.
On ordinary MCP client shutdown, the package printed its own traceback ending
in:

```text
ValueError: I/O operation on closed file.
```

The benchmark protocol includes arm C only if it installs **and runs cleanly**
in at most 15 minutes. Because this reproducible shutdown failure violates that
gate, the arm is skipped. It is not charged against task accuracy.

## Generated assets

- `docs/assets/benchmark-token-hero.{png,svg}`
- `docs/assets/benchmark-token-modes.{png,svg}`
- `docs/assets/benchmark-tool-calls.{png,svg}`
- `docs/assets/benchmark-index-time.{png,svg}`
- `docs/assets/benchmark-audit-cost.{png,svg}`

The dollar callout states that Codex CLI did not report a dollar field. The
wall-time metric remains in `accuracy.csv` but is not substituted for the
required cost callout.

## Criterion conclusion

S1 passes. S5's equal-or-better-accuracy clause passes, but its ≥10× token
reduction clause fails in both measurement modes. The criterion therefore
fails as frozen. No chart, README copy, or release evidence may describe S5 as
met.

## Candidate verification

The final pre-freeze engineering pass found one semantic issue in the numeric
parser optimization: an integral decimal beyond IEEE-754's exact integer range
could round before conversion to `int`. The fast path now delegates that
uncommon branch to the canonical `Decimal` parser, and a regression covers
`9007199254740993.0` plus ordinary decimal, exponent, overflow, and underflow
forms.

Fresh checks on the remediated candidate:

```text
uv run pytest -q --cov=excel_lsp.core --cov-report=term --cov-report=json:coverage.json
  2102 passed, 3 deselected; total core branch coverage 89.63% (gate: 85%)
uv run pytest -m live tests/live/test_p8_evidence.py -q
  2 passed
uv run ruff check .
  passed
uv run ruff format --check .
  111 files already formatted
uv run pyright
  0 errors, 0 warnings, 0 informations
uv run python tests/fixtures/generate.py
  generated F01-F21 successfully
```

A separate three-repetition timing verification on the final code, written to
stdout rather than over the committed measurement rows, produced a 50,000-row
cold median of 9.431842 seconds and incremental median of 0.075041 seconds.
S1 therefore remained green after remediation.
