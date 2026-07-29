# Examples

This directory contains verified configuration examples for Excel LSP's v0.1.0
stdio MCP server. Clean wheel, `uvx`, MCP initialization, and Codex registration
results are recorded in the [install evidence](../docs/evidence/fresh-install.md).

## Codex configuration

[`codex.config.toml`](codex.config.toml) is the native Codex configuration:

```toml
[mcp_servers.excel-lsp]
command = "uvx"
args = ["--from", "git+https://github.com/Brody900/excel-lsp@main", "excel-lsp", "serve"]
```

The equivalent Codex CLI registration is:

```console
codex mcp add excel-lsp -- uvx --from git+https://github.com/Brody900/excel-lsp@main excel-lsp serve
```

Codex CLI 0.144.5 accepted this exact fallback command in an isolated
configuration home, returned the expected stdio command and arguments, and
removed it cleanly. The server also completed real MCP initialization from the
same public source.

## Generic MCP JSON

[`mcp.json`](mcp.json) contains the same command in the common `mcpServers`
JSON shape. It is for MCP clients that consume that format; it is not Codex's
native configuration file.

After PyPI publication is visible, the shorter arguments are
`["excel-lsp", "serve"]`. The [README](../README.md) contains copy-paste
commands for both paths.
