# Registry submission packet

Prepared **2026-07-29** for Excel LSP v0.1.0. The public repository is
<https://github.com/Brody900/excel-lsp>. Submissions that require a maintainer
account, a published PyPI artifact, a paid checkout, or an authenticated CLI
remain explicit follow-ups rather than being represented as completed.

## Shared listing copy

**Name:** Excel LSP

**Short description:** Semantic workbook index, dependency navigation,
diagnostics, and surgical Excel edits for MCP agents.

**Long description:** Excel LSP gives MCP-compatible agents a persistent
semantic view of modern Excel OOXML workbooks. Its 14 bounded tools map sheets
and regions, navigate formula precedents and dependents, explain formulas,
surface diagnostics, profile columns, and make two narrow surgical edits while
preserving untouched ZIP parts byte-for-byte. It runs locally over stdio, makes
no runtime network requests, and supports optional `EXCEL_LSP_ROOT` path
confinement.

**Repository:** <https://github.com/Brody900/excel-lsp>

**Install command:** `uvx excel-lsp serve`

**Fallback command:**
`uvx --from git+https://github.com/Brody900/excel-lsp@v0.1.0 excel-lsp serve`

**Categories:** Microsoft Excel, spreadsheets, developer tools, local files,
data analysis.

**License:** MIT.

## Official MCP Registry

The ready-to-publish [`server.json`](../server.json) uses the official
2025-12-11 schema, the verified GitHub repository ID, a PyPI package entry, the
`uvx` runtime hint, and stdio transport. README contains the matching ownership
token:

```text
mcp-name: io.github.Brody900/excel-lsp
```

After the PyPI package is public, the maintainer follow-up is:

```console
mcp-publisher login github
mcp-publisher validate
mcp-publisher publish
```

The official registry is still preview software. Publishing before PyPI would
fail its package ownership check, so the metadata is prepared but submission is
sequenced after package publication.

## Smithery

Smithery's current local-stdio route requires an MCPB bundle and an authenticated
namespace. The exact listing copy is the shared copy above. Maintainer steps:

1. Package the released `uvx excel-lsp serve` launcher as a `.mcpb` bundle.
2. Authenticate with `smithery auth login` and select the intended namespace.
3. Run `smithery mcp publish ./excel-lsp-0.1.0.mcpb -n <namespace>/excel-lsp`.
4. Confirm that the generated installation command launches the released PyPI
   package, then link the evidence index and security policy.

No Smithery credential or namespace is available in this workspace, so no
account-scoped publication is claimed.

## mcp.so

Use <https://mcp.so/submit?type=server> with the shared copy and repository URL.
As accessed on 2026-07-29, submission requires a paid one-time checkout. No
purchase was authorized, so the packet is ready and payment is a user follow-up.

## PulseMCP

Use <https://www.pulsemcp.com/submit> with the shared copy, repository URL,
local stdio transport, and install command. PulseMCP also ingests the official
MCP Registry, so an accepted official listing may be discovered automatically.
The submission form requires a browser/account flow not available as a
non-interactive repository artifact; no submission is claimed.

## Publication state

| Directory | Prepared | Submitted | Remaining external action |
|---|---:|---:|---|
| Official MCP Registry | Yes | No | Publish PyPI package, authenticate, publish `server.json` |
| Smithery | Yes | No | Choose namespace, create MCPB, authenticate |
| mcp.so | Yes | No | Authorize and complete paid checkout |
| PulseMCP | Yes | No | Complete maintainer submission form or await registry ingestion |
