# P9 fresh-install evidence

Executed **2026-07-29** from the P9 release candidate on Microsoft Windows 11
Home 64-bit, version 10.0.26200 (build 26200). Source base commit:
`4e8e0852a69e3786020356ef66cb90bcff0916a8`. Tool versions: Python 3.12.11,
uv 0.9.30, Codex CLI 0.144.5.

All temporary environments had unique directories below the operating-system
temporary root. Paths below use `<TEMP>` and `<WORKSPACE>` so the evidence does
not publish a local username or workstation layout.

## Build and metadata

```console
uv build
uv run --with twine twine check dist/*
```

Both `excel_lsp-0.1.0-py3-none-any.whl` and
`excel_lsp-0.1.0.tar.gz` built successfully; Twine passed both. The wheel embeds
the full README, MIT license, Python requirement, public project URLs, and the
`excel-lsp` console entry point. Wheel SHA-256:
`16400354213378e339170a2e0f5c7da214d2c3c5caae0afc9fdb512c0a30665d`.

The official MCP Registry 2025-12-11 JSON Schema validated
[`server.json`](../../server.json), and its PyPI ownership name matches the
README's `mcp-name` token.

## Wheel install

The wheel was installed into a new Python 3.12 virtual environment with no
editable source path:

```console
uv venv <TEMP>/venv --python 3.12
uv pip install --python <TEMP>/venv/Scripts/python.exe \
  <WORKSPACE>/dist/excel_lsp-0.1.0-py3-none-any.whl
<TEMP>/venv/Scripts/excel-lsp.exe --version
<TEMP>/venv/Scripts/excel-lsp.exe map \
  <WORKSPACE>/tests/fixtures/generated/cross_sheet_model.xlsx
<TEMP>/venv/Scripts/python.exe tests/release/probe_mcp.py \
  <TEMP>/venv/Scripts/excel-lsp.exe \
  <WORKSPACE>/tests/fixtures/generated/cross_sheet_model.xlsx
```

Observed: version `0.1.0`; map completed with three sheets; a real MCP client
initialized the stdio server, listed exactly 14 tools, received instructions,
opened the workbook, and received implementation name `excel-lsp` and version
`0.1.0`.

The first isolated probe reproduced a release-only defect: FastMCP substituted
its SDK version for an unset implementation version. The candidate now sets the
low-level implementation version from `excel_lsp.__version__`, and both the MCP
conformance test and release probe protect that behavior.

## uvx install

The freshly built wheel also passed the runner path without installation into
the repository environment:

```console
uvx --from <WORKSPACE>/dist/excel_lsp-0.1.0-py3-none-any.whl \
  excel-lsp --version
uvx --from <WORKSPACE>/dist/excel_lsp-0.1.0-py3-none-any.whl \
  excel-lsp map <WORKSPACE>/tests/fixtures/generated/cross_sheet_model.xlsx
uv run python tests/release/probe_mcp.py <UVX_EXECUTABLE> \
  <WORKSPACE>/tests/fixtures/generated/cross_sheet_model.xlsx \
  --uvx-package <WORKSPACE>/dist/excel_lsp-0.1.0-py3-none-any.whl
```

Observed: version `0.1.0`, map pass, MCP initialization pass, 14 tools, and a
successful `open_workbook` call.

## Git fallback

The public repository fallback was exercised against the immutable public P8
commit—the newest public commit at test time—rather than a local checkout:

```console
uvx --from \
  git+https://github.com/Brody900/excel-lsp@4e8e0852a69e3786020356ef66cb90bcff0916a8 \
  excel-lsp --version
uvx --from \
  git+https://github.com/Brody900/excel-lsp@4e8e0852a69e3786020356ef66cb90bcff0916a8 \
  excel-lsp map <WORKSPACE>/tests/fixtures/generated/cross_sheet_model.xlsx
```

The public clone built, reported package version `0.1.0`, mapped the fixture,
initialized over stdio, listed 14 tools, and opened the workbook. The P8 server
reported the SDK version during initialization; the wheel evidence above proves
that the P9 candidate fixes that separately reproduced metadata defect. After
release, users can pin `@v0.1.0`; the README's `@main` command remains
copy-pasteable before the tag exists. The exact current README fallback was
rerun after the first formal review: public `@main` again initialized through a
real MCP client, listed 14 tools, and opened F03.

## Codex MCP registration

Fresh CLI syntax is captured in [`codex-mcp-help.txt`](codex-mcp-help.txt).
With an isolated `CODEX_HOME`, the exact quickstart command succeeded:

```console
codex mcp add excel-lsp -- uvx --from git+https://github.com/Brody900/excel-lsp@main excel-lsp serve
codex mcp get excel-lsp
codex mcp remove excel-lsp
```

`get` returned `transport: stdio`, `command: uvx`, and the exact four fallback
arguments; removal succeeded. Codex declined only to create its own helper
aliases under the temporary directory, an intentional host safety rule
unrelated to MCP registration. The real user configuration was untouched.

## Codex TOML configuration

[`examples/codex.config.toml`](../../examples/codex.config.toml) has the same
command and arguments as the CLI registration. The generic
[`examples/mcp.json`](../../examples/mcp.json) uses the equivalent common MCP
JSON shape. `tests/unit/test_readme_contract.py` parses both and requires their
server commands to remain equivalent.

## Runtime network audit

```console
uv run pytest -q \
  tests/mcp/test_conformance.py::test_server_operates_with_network_denied
```

Result: `1 passed`. The test denies socket connections and DNS lookup, then
indexes a workbook and executes representative list/read operations through
the server service. OOXML parsers also set lxml's `no_network=True`.

## PyPI and fallback state

The official `https://pypi.org/pypi/excel-lsp/json` endpoint returned 404 on
2026-07-29, so the project name remained unclaimed. No recognized PyPI token
environment variable was present. The release workflow builds and verifies
artifacts on a tag; its trusted-publishing job runs only when the repository
variable `PYPI_PUBLISH_ENABLED` is `true`. This avoids a false failing release
when no publisher has been configured.

## S7 conclusion

**Pass.** A clean wheel and local-wheel `uvx` path both initialize the released
server surface, and the documented public-git fallback independently builds and
runs from GitHub. PyPI publication is desirable but is not required by the
frozen criterion when a working git fallback is documented.
