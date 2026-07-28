# Phase 7 — MCP server and CLI evidence

Status: verified; combined R-mech/R-test review approved.

Base commit before P7: `701b6832b5138010aab8063b9dade7371154cb73`.
Environment: Windows, Python 3.12.11, uv 0.9.30, MCP SDK 1.28.1,
`regex` 2026.7.10.

## Package and transport boundary

The candidate adds one shared `ToolService` used by both the FastMCP stdio
server and Typer CLI. Core remains importable and usable independently; only
the lifecycle gained an optional callback for real in-flight per-sheet progress.

- `src/excel_lsp/server/app.py`: exactly 14 FastMCP tools, initialization
  instructions, annotations, progress bridging, and sanitized unexpected errors.
- `src/excel_lsp/server/service.py`: freshness, queries, graph projection,
  editing, response shaping, value caps, and shared CLI/MCP behavior.
- `src/excel_lsp/server/cursors.py`: opaque parameter- and generation-bound
  pagination.
- `src/excel_lsp/server/security.py`: optional realpath allowlist.
- `src/excel_lsp/cli/app.py`: serve, map, trace, path, diag, find, schema,
  graph/Mermaid, and the P8 benchmark entry point.

The production source never loads and saves an existing workbook through
openpyxl. Both exposed writes call the verified P6 surgical editor.

## Tool inventory and schemas

`tools/list` returns exactly 14 tools: 12 reads and 2 destructive writes.
Every read has `readOnlyHint=true` and `openWorldHint=false`;
`write_cells` and `set_column_formula` have `readOnlyHint=false` and
`destructiveHint=true`.

The exact generated input schemas and annotations are committed in
[`p7-tool-schemas.json`](p7-tool-schemas.json). The subprocess conformance test
recreates the dictionary from `tools/list` and requires exact equality. The
reserved JSON names `in` and `max` are also exercised through a real client
call. Each tool has a worked request/response in
[`docs/tool-reference.md`](../tool-reference.md).

## Freshness, progress, and cursors

Every service method realpath-resolves the workbook and calls
`ensure_fresh` or `index_workbook`. A cold index reports each completed
sheet from inside the index transaction through the MCP progress token. A
no-op open still reports the existing sheet inventory when a token is present.

`read_range` pages in deterministic row-major order, includes explicit page
start/end coordinates, and caps each page at 200 cells. Its opaque base64 cursor
binds the tool, normalized query hash, offset, and index generation. The stdio
test obtains a cursor, advances it, performs a surgical write, and verifies the
old cursor returns `E_STALE_CURSOR`. It also changes the workbook outside the
server and proves the next read tool returns `reindexed=true`.

## Response caps

The service measures the same pretty JSON representation emitted by FastMCP and
reduces bounded list detail deterministically until it is at most 8,000
characters. The transport applies the same cap to canonical error envelopes,
truncating echoed input and details while preserving the stable error code.
The subprocess test independently measures both unstructured MCP text and
pretty-serialized structured data. Focused regressions cover:

- F20 stress-map degradation;
- a 181-column region, where `sample_rows=0` and no sample values escape;
- row-boundary pagination;
- a maximum 500-cell write that executes all edits while returning a bounded,
  explicitly truncated result with `resultsTotal=500`;
- a 12,000-character unknown symbol in both direct invocation and stdio;
- cell strings too large for one response.

## Regex guard

`find` limits patterns to 256 characters, subjects to 1,000, snippets to 80,
and the whole scan to two seconds. The production dependency is the
timeout-capable `regex` engine. A regression authors a 1,000-character
nonmatching subject and proves `^(a|aa)+b$` is interrupted within the safety
window with partial output and `W_REGEX_TIMEOUT`.

## Path confinement

If `EXCEL_LSP_ROOT` is unset, access follows the process's local filesystem
permissions. When set, it is an `os.pathsep`-separated list of existing
directories. Requested paths and roots are resolved before a case-normalized
common-path comparison. Unit coverage proves allowed access, outside denial,
and symlink escape denial. The subprocess test runs with the allowlist enabled
and receives canonical `E_PATH_DENIED` for an outside workbook.

## CLI evidence

Typer tests verify the complete command inventory, JSON equivalence for map,
trace, path, diagnostics, find, and schema, exact direction validation, and
Mermaid flowchart output. `excel-lsp serve` is the stdio entry point exercised
by T6. The `bench` command is registered now and delegates to the P8 benchmark
runner; P8 owns that harness and its measured artifacts.

## Fresh verification

The first combined P7 reviewer returned `R-mech: REVISE` and
`R-test: REVISE` on the original frozen tree. Its four major reproductions are
now permanent focused and stdio regressions:

- expected error envelopes pass through the same 8,000-character fitter;
- cell-symbol queries stream the complete indexed domain, retain only the
  deterministic first 100 matches, and report the exact full match count;
- value searches conservatively expose cached-value staleness for their search
  scope;
- profile `count` measures resolved range positions while `nonnull` measures
  persisted normalized values, including direct and semantic sparse columns.

The reviewer also found one current README comparison row that still called
the committed P6 milestone a candidate; README and SECURITY wording now agree
with the historical P6 approval and ledger.

The second combined review reproduced a behaviorally distinct serialization
boundary: a one-cell value containing 4,000 newlines fit the code-point guard
but expanded beyond 8,000 characters as pretty JSON. `read_range` now
binary-searches the largest deterministic prefix that fits the actual emitted
representation and reports `valueTruncated=true`. This value-level reduction
does not falsify pagination: a complete one-cell request still reports
`truncated=false` and `cursor=null`. Both the direct service test and real stdio
conformance exercise that exact escaped-string case.

| Command | Result |
|---|---|
| `uv run pytest tests/unit/test_server_service.py tests/mcp/test_conformance.py tests/unit/test_cli.py tests/unit/test_readme_contract.py -q` | 34 passed after second formal-review remediation, 9.59 s |
| `uv run pytest` | 2,055 passed, 1 live deselected, 400.05 s |
| `uv run pytest --cov=excel_lsp.core --cov-branch --cov-report=term-missing` | 2,055 passed, 1 live deselected, 89.67% branch coverage, 785.74 s |
| `uv run ruff check .` | passed |
| `uv run ruff format --check .` | 94 files already formatted |
| `uv run pyright` | 0 errors, 0 warnings, 0 informations |
| `uv lock --check` | 70 packages resolved; lock current |
| `uv build` | wheel and sdist built |
| `uv run python tests/fixtures/generate.py` | 20 fixture IDs regenerated; no tracked drift |
| `git diff --check` | passed |

The full suite and coverage rows were run after all production-source changes,
including the 500-cell response-cap regression. The only failure in the first
exact-candidate rerun was the claims-matrix validator rejecting the legitimate
pre-review `Candidate P7` status; after the validator and status documentation
were aligned, all 12 README contract tests and the complete 2,050-test suite
passed. Final accounting-sensitive checks were rerun immediately before the
fingerprint freeze.

## Formal phase gate

The third combined review returned clean `R-mech: APPROVE` and
`R-test: APPROVE` verdicts on one unchanged fingerprint:

- base: `701b6832b5138010aab8063b9dade7371154cb73`;
- staged tree: `2e7d105387e44a4acbdb674f6b73121e1d2f0d2f`;
- staged binary-diff hash: `b3cb3077e40963b328d82fc4fc7095ab1a7a5798`;
- scope: 27 files, 3,127 insertions, 180 deletions.

The reviewer reported no critical, major, or minor findings, independently
passed the 34-test P7/CLI/accounting slice and static checks, and confirmed
entry and exit fingerprints were identical with no unstaged or untracked
files. Its separate escaped-value probe verified both paginated and terminal
single-cell semantics at 7,999–8,000 serialized characters. The P7 formal gate
is closed.

## Phase boundary

P7 does not claim the P8 live-Excel protocol, six-task benchmark harness,
headless-Codex evaluations, charts, or benchmark conclusions. It does not claim
the P9 clean `uvx` install, public package, or registry submissions. Those
remain explicit later gates.
