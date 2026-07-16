# Phase 0 reconnaissance evidence

Recorded on 2026-07-15 (America/Los_Angeles). Secrets and credential values are
intentionally omitted.

| Capability | Observed result |
|---|---|
| Host | Windows 11 Home, x64, build family 26200 |
| Python | 3.11, 3.12, 3.13, and 3.14 interpreters available; project uv environment uses 3.12.11 |
| uv | 0.9.30 |
| Git | 2.53.0; repository branch normalized to `main` |
| GitHub | GitHub CLI 2.88.1 authenticated with repository and workflow scopes |
| Codex | CLI 0.144.1; authenticated headless JSON ping returned exactly `PONG` |
| Excel | Desktop Excel 16.0 build 19530, x64; hidden COM launch and clean quit succeeded |
| SQLite | 3.47.1; `CREATE VIRTUAL TABLE ... USING rtree` plus insert/query succeeded |
| PyPI name | `excel-lsp` JSON and project endpoints returned 404, so the preferred name was available at check time |

## Deterministic dependency choices

- `openpyxl==3.1.5`
- `tiktoken==0.13.0`
- 70 packages resolved in `uv.lock`; `uv sync --all-extras --dev --locked`
  audited the environment successfully.

The two exact pins were the current stable PyPI releases observed during P0 and
both support Python 3.11. Other dependencies use compatible ranges; `uv.lock`
captures the complete cross-platform resolution.

## Headless Codex verification

The successful probe used an ephemeral, read-only, JSON run with the normal
authenticated configuration and no tools. The final agent message was `PONG`.
The same probe with `--ignore-user-config` failed with HTTP 401, so the Phase 8
eval harness will retain authentication and isolate MCP tools/rules with
Codex-native profiles or explicit configuration instead. No Claude or Anthropic
service is part of the benchmark plan.

## Excel and VBA verification

The local-only `f16_source.xlsm` was inspected as OOXML and with `olevba`. Its
VBA project contains the intended `Stamp` macro. A hidden Excel COM session
opened the workbook read-only, executed the macro, observed `Z1 = 42`, and
closed without saving. The source SHA-256 before and after was identical:

```text
cf1c3016d5409905a33fbdbceaf2612f060dbba13274438a0936fbcca1aadb46
```

Only `tests/fixtures/assets/vbaProject.bin` is tracked. Its SHA-256 is:

```text
be05aafbb31d2de0ffd686c9cae71b97a2596132ed4443d9d656558d7089ccb1
```

## Phase 0 self-check

Fresh orchestrator-run results:

| Check | Result |
|---|---|
| `uv lock --check` | passed |
| `uv sync --all-extras --dev --locked` | passed |
| `uv run ruff check .` | passed |
| `uv run ruff format --check .` | passed; 8 files already formatted |
| `uv run pyright` | passed; 0 errors, warnings, or information messages |
| `uv run pytest --cov=excel_lsp.core --cov-fail-under=85` | 4 passed; scaffold core coverage 100% |
| `uv run python tests/fixtures/generate.py` | passed; Phase 0 skeleton executed |
| `uv run excel-lsp --version` | `0.1.0` |
| `uv build` | built sdist and wheel successfully |

The scaffold deliberately implements no workbook behavior; parser and fixture
functionality begins at Phase 1.
