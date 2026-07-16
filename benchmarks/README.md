# Benchmarks

P2 records one narrow, deterministic measurement:
[`results/map-budgets.json`](results/map-budgets.json) contains the normalized
F03 and F20 workbook-map character counts and `o200k_base` token counts. The
golden-map integration test recomputes those values with pinned
`tiktoken==0.13.0`. This is response-budget evidence, not a comparative
performance, accuracy, cost, or token-reduction benchmark.

P8 adds the reproducible benchmark runners, deterministic task definitions,
exact `ANSWER:` checkers, scripted replays, both headless-Codex repetitions,
timing series, raw results, and chart-generation scripts. Until then, the README
makes no comparative numerical claim.

Raw measurements belong in `results/`; deterministic task definitions and
checkers belong in `tasks/`. The methodology and exact future filenames are
tracked in
[`docs/evidence/readme-claims-to-artifacts.md`](../docs/evidence/readme-claims-to-artifacts.md).
