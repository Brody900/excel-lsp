# Security policy

Excel LSP is a local-file semantic index for Excel OOXML workbooks. It is not a
hosted service, and the v0.1.0 runtime is designed to make no network requests.
Because an MCP server can be given paths and can eventually write workbooks,
the filesystem boundary is part of the security model rather than an implicit
promise.

## Release status

Version 0.1.0 includes the parser, index, regions, symbols, formula analysis,
graph, diagnostics, surgical editor, bounded stdio MCP/CLI surface, realpath
confinement, live Excel evidence, benchmarks, and clean-install evidence.

## Supported files and trust boundary

The release target supports `.xlsx`, `.xlsm`, `.xltx`, and `.xltm` files on the
local filesystem. Legacy binary `.xls` and `.xlsb` files are not supported.
Workbook XML, formulas, strings, relationships, and embedded package metadata
are untrusted input. Parsing failures are shaped into documented errors rather
than exposing internal tracebacks through MCP.

The SQLite sidecar is a derived index. It is safe to delete when Excel LSP is
not using it; the source workbook remains authoritative. Do not commit index
sidecars or use them as backups.

## Path access

The MCP server supports optional confinement through `EXCEL_LSP_ROOT`. The value is an
`os.pathsep`-separated allowlist of directories. Every workbook path is
realpath-resolved, including symlinks, and must remain within one of those
directories or return `E_PATH_DENIED`.

The default is deliberately explicit: if `EXCEL_LSP_ROOT` is unset, local-path
access is unrestricted to the permissions of the process running the MCP
server. Users who connect an autonomous agent should set the allowlist to the
smallest workbook directory it needs.

## Write scope and workbook fidelity

Excel LSP exposes exactly two edit tools:
`write_cells` and `set_column_formula`. They use surgical OOXML patches rather
than loading and saving an existing workbook through openpyxl. A successful
edit may modify the targeted worksheet XML and required calculation metadata,
including deliberate calc-chain handling. Every other ZIP part must remain
byte-identical.

Writes refuse an Excel-open workbook when its lockfile is present, reject a
workbook changed since indexing or immediately after replacement, reject
unsupported values, and refuse cells
inside multi-cell array formulas. The release gate includes deterministic
part-diff tests on macro, chart, and image fixtures plus a live Excel pass. The
P6 verified evidence is in [`docs/evidence/p6-editor.md`](docs/evidence/p6-editor.md);
P7's stdio conformance suite covers their structured tool surface and path
confinement, while P8 owns the complete live protocol.

## Runtime data handling

- Excel LSP has no telemetry.
- The release runtime is designed to perform no network access.
- Tool responses are capped and are not intended to dump whole workbooks.
- Workbook data and index sidecars remain on the local machine unless the MCP
  client or another program transmits them.
- Logs and bug reports must not include private workbook contents, credentials,
  signed URLs, or environment-variable values.

## Dependency and release practices

Dependencies are locked in `uv.lock`. CI exercises Python 3.11–3.13 on Linux,
macOS, and Windows. Release artifacts are built from the reviewed source, and
the clean-install evidence identifies the tested commit, Python version,
operating system, Codex CLI version, and artifact hashes.

## Reporting a vulnerability

Do not disclose a suspected vulnerability in a public issue with an exploit or
private workbook attached. Use GitHub's private
[security-advisory form](https://github.com/Brody900/excel-lsp/security/advisories/new).

Include the affected commit or version, platform, minimal reproduction using
synthetic data, impact, and whether the issue permits reading or modifying paths
outside the intended scope. Remove secrets and personal workbook data first.

## Supported versions

| Version | Supported |
|---|---:|
| 0.1.x | Yes |
| Earlier or untagged builds | No |
