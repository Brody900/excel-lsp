# Excel LSP — working conventions

Compatibility file retained because the frozen handoff names it in the
Definition of Done. Codex is the project orchestrator; use `AGENTS.md` for the
active repository instructions and treat every Claude reference in the handoff
as Codex.

One-liner: LSP-style semantic index + MCP server for Excel workbooks.

Layout: `src/excel_lsp/{core,server,cli}`; `tests/{unit,property,golden,oracle,fixtures,mcp}`; `benchmarks/`; `docs/`. Plans: `PLAN.md`. Decisions: `docs/agent-log.md`. Review ledger: `PLAN.md` table.

Frozen decisions: see `HANDOFF.md` §0.4. Do not renegotiate them in code.

Before touching a module, read its HANDOFF section:

- parser/loader → §5.1
- regions → §5.2
- classify_ref → §5.3
- R1C1/blocks → §5.4
- graph → §5.5
- diagnostics → §5.6
- editor → §5.7
- lifecycle → §5.8
- tools/server → §6
- tests → §8
- benchmarks → §9

Hard rules:

- NEVER load-and-save a workbook through openpyxl in `src/` (HANDOFF §5.7). Edits are surgical.
- No bulk data in tool responses; caps per HANDOFF §6.1.
- Value normalization goes through the one shared function (§5.1).

Commits: conventional (`feat:`/`fix:`/`test:`/`docs:`/`chore:`), per milestone, on `main`, no force-push.

Tests: `uv run pytest`

Lint/type: `uv run ruff check . && uv run pyright`

Fixtures: `uv run python tests/fixtures/generate.py`
