# Examples

This directory contains configuration examples for Excel LSP's planned v0.1.0
stdio MCP server. They are committed early so the P2 repository review can
check the public quickstart and configuration shape. They are not release proof:
the server arrives in P7, and P9 must execute each path from a clean environment
before the pre-release warning is removed from the README.

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

Both examples currently describe a future published command. Until P7 and P9
finish, use the repository's locked development environment and treat
[`README.md`](../README.md) as pre-release documentation. The exact install and
configuration evidence required for publication is tracked in the
[README claims-to-artifacts plan](../docs/evidence/readme-claims-to-artifacts.md).
