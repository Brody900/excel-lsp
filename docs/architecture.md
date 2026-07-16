# Architecture

Excel LSP is organized as three strictly separated Python layers:

- `excel_lsp.core`: pure workbook parsing, indexing, graph, diagnostics, and
  surgical editing code with no MCP dependency.
- `excel_lsp.server`: MCP transport, schemas, response limits, and pagination.
- `excel_lsp.cli`: the Typer command-line interface over the same core library.

Phase 0 establishes only these package boundaries. Implementations are added in
the gated phases described by `PLAN.md`.
