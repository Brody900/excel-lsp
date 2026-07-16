# Security policy

Excel LSP is a local-file semantic index for Excel OOXML workbooks. It is not a
hosted service, and the v0.1.0 runtime is designed to make no network requests.
Because an MCP server can be given paths and can eventually write workbooks,
the filesystem boundary is part of the security model rather than an implicit
promise.

## Development status

This repository is pre-release. The streaming parser, derived SQLite index, and
freshness lifecycle are verified. Sparse regions, stable symbols, and the
compact workbook map are implemented and tested in the active P2 worktree but
remain gate-pending. The MCP server, path-confinement boundary, and two write
tools are gated for P6-P7 and must not be described as available until their
tests and evidence are committed.

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

P7 will add optional confinement through `EXCEL_LSP_ROOT`. The value is an
`os.pathsep`-separated allowlist of directories. Every workbook path will be
realpath-resolved, including symlinks, and must remain within one of those
directories or return `E_PATH_DENIED`.

The default is deliberately explicit: if `EXCEL_LSP_ROOT` is unset, local-path
access is unrestricted to the permissions of the process running the MCP
server. Users who connect an autonomous agent should set the allowlist to the
smallest workbook directory it needs.

## Write scope and workbook fidelity

P6 will implement exactly two write tools: `write_cells` and
`set_column_formula`. They will use surgical OOXML patches rather than loading
and saving an existing workbook through openpyxl. A successful edit may modify
the targeted worksheet XML and required calculation metadata, including
deliberate calc-chain handling. Every other ZIP part must remain byte-identical.

Writes will refuse an Excel-open workbook when its lockfile is present, reject
a workbook changed since indexing, reject unsupported values, and refuse cells
inside multi-cell array formulas. The release gate includes deterministic
part-diff tests on macro, chart, and image fixtures plus a live Excel pass. None
of those statements should be read as a claim that editing is implemented in
the current P2 codebase.

## Runtime data handling

- Excel LSP has no telemetry.
- The release runtime is designed to perform no network access.
- Tool responses are capped and are not intended to dump whole workbooks.
- Workbook data and index sidecars remain on the local machine unless the MCP
  client or another program transmits them.
- Logs and bug reports must not include private workbook contents, credentials,
  signed URLs, or environment-variable values.

## Dependency and release practices

Dependencies are locked in `uv.lock`. CI will exercise the supported Python and
operating-system matrix before release. Release artifacts must be built from a
clean checkout, and the clean-install evidence must identify the tested commit,
Python version, operating system, and Codex CLI version.

## Reporting a vulnerability

Do not disclose a suspected vulnerability in a public issue with an exploit or
private workbook attached. Once the public GitHub repository enables private
vulnerability reporting, use its **Security → Report a vulnerability** flow.
Before that P9 publication step, report the issue privately to the repository
owner through the same private channel used to receive this source tree. This
section will be updated with the final private-advisory URL at release time.

Include the affected commit or version, platform, minimal reproduction using
synthetic data, impact, and whether the issue permits reading or modifying paths
outside the intended scope. Remove secrets and personal workbook data first.

## Supported versions

No public version is supported yet. After v0.1.0, the latest tagged minor line
will receive security fixes; the precise support window will be recorded here
before the first release.
