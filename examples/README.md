# Examples

This directory contains configuration examples for Excel LSP's v0.1.0 stdio
MCP server. The server is verified in P7, but these published
`uvx` commands are not release proof until P9 executes each path from a clean
environment.

## Codex configuration

[`codex.config.toml`](codex.config.toml) is the native Codex configuration:

```toml
[mcp_servers.excel-lsp]
command = "uvx"
args = ["excel-lsp", "serve"]
```

The equivalent Codex CLI registration planned for the release quickstart is:

```console
codex mcp add excel-lsp -- uvx excel-lsp serve
```

P9 will re-check this syntax against the then-current Codex CLI, record the
version and output, initialize the server, and remove the configuration again.

## Generic MCP JSON

[`mcp.json`](mcp.json) contains the same command in the common `mcpServers`
JSON shape. It is for MCP clients that consume that format; it is not Codex's
native configuration file.

Both examples currently describe a future published command. Until P9
finishes, use `uv run excel-lsp serve` in the repository's locked development
environment and treat
[`README.md`](../README.md) as pre-release documentation. The exact install and
configuration evidence required for publication is tracked in the
[README claims-to-artifacts plan](../docs/evidence/readme-claims-to-artifacts.md).
