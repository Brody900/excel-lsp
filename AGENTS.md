# Excel LSP — Codex working conventions

One-liner: LSP-style semantic index + MCP server for Excel workbooks.

`HANDOFF.md` v1.1 is the authoritative specification. Work through P0–P9 in
order and do not skip phase gates. `PLAN.md` owns phase/review accounting;
`docs/agent-log.md` is an append-only decision log.

## Layout

- `src/excel_lsp/core/`: OOXML parser, index, graph, diagnostics, editor.
- `src/excel_lsp/server/`: MCP stdio server, schemas, caps, pagination.
- `src/excel_lsp/cli/`: Typer debugging and benchmark CLI.
- `tests/`: unit, property, golden, oracle, fixtures, and MCP conformance.
- `benchmarks/`, `docs/`, and `examples/`: release evidence and public copy.

Before touching a module, read its `HANDOFF.md` section:

- parser/loader → §5.1; regions → §5.2; reference extraction → §5.3
- R1C1/blocks → §5.4; graph → §5.5; diagnostics → §5.6
- editor → §5.7; lifecycle → §5.8; tools/server → §6
- tests → §8; benchmarks → §9

## Frozen rules

- Never load and save an existing workbook through openpyxl in `src/`.
  Production edits are surgical OOXML patches and must satisfy I18.
- No tool returns bulk data: at most 200 raw values and 8,000 serialized
  characters, with generation-bound pagination where specified.
- All public values use the one shared normalization function from §5.1.
- Preserve the frozen symbol IDs, response caps, license, package name, review
  budgets, phase gates, and Definition of Done.
- Claude references in the historical handoff mean Codex. Do not invoke Claude
  or Anthropic services. Verify current Codex CLI syntax before headless evals.

Use conventional milestone commits on `main`; never force-push. Preserve user
work and keep unrelated changes out of commits.

## Standard checks

```text
uv run pytest
uv run ruff check .
uv run ruff format --check .
uv run pyright
uv run python tests/fixtures/generate.py
```

Run targeted checks first, then the full suite. Never claim a criterion without
fresh command output and a committed evidence artifact.
