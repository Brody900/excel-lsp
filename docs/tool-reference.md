# Tool reference

This document records the frozen v0.1.0 MCP contracts before their P7
implementation. It is not a claim that the server is available in the active
P2 worktree. P7 must add one fully worked request/response example, happy-path
and error-path conformance, and exact generated schemas under every heading.

The first 12 tools are reads and will declare `readOnlyHint: true` and
`openWorldHint: false`. `write_cells` and `set_column_formula` are destructive
writes and will declare `readOnlyHint: false` and `destructiveHint: true`.
Every call performs a freshness check. No response may exceed 8,000 serialized
characters, and no response may contain more than 200 raw cell values.

## `open_workbook`

Input: `path`.

Output: the compact workbook map: sheet dimensions and visibility, bounded
regions and column summaries, names, external links, diagnostics counts, VBA
presence, staleness, and navigation hints. It returns no raw body-cell values.
P2 implements the core map projection; P7 adds this MCP entry point and
per-sheet progress notifications.

## `refresh`

Input: `path`, `recalculated: bool = false`.

Output: the refreshed workbook map and `reindexedSheets`. When `recalculated`
is true, P6 staleness may be cleared according to the editor contract. P7 must
also report indexing progress when the client supplies a progress token.

## `list_symbols`

Input: `path`, `query: str = ""`, `kinds: list[str] = all`.

Output: matching stable symbol IDs with one-line descriptors. The frozen kinds
cover sheets, regions, columns, defined names, formula blocks, and cells.

## `get_region_schema`

Input: `path`, `region_id`.

Output: headers, inferred dtypes, non-null and distinct counts, validation
summaries, formula-block summaries, confidence, and bounded sample rows. Sample
rows use `max(0, min(3, 180 // ncols))`; when zero, the response directs the
caller to `read_range`.

## `read_range`

Input: `path`, `ref`, optional `cursor`, `max_cells: int = 200`.

Output: a two-dimensional value page plus `truncated`, `cursor`, and `stale`.
This is the only tool that returns values in quantity. Cursors bind tool
parameters, offset, and index generation; mutation before the next page yields
`E_STALE_CURSOR`.

## `find`

Input: `path`, `pattern`, `in` selected from values, headers, formulas, and
names, optional `sheet`, `max: int = 50`.

Output: bounded `{ref, kind, snippet}` matches with snippets no longer than 80
characters. P7 must enforce the regex deadline and return `W_REGEX_TIMEOUT`
rather than allowing unbounded expression work.

## `trace_precedents`

Input: `path`, `ref_or_symbol`, `depth: int = 2` capped at 8, and
`max_nodes: int = 200`.

Output: a bounded tree of upstream symbols or references and edge reasons,
including a truncation flag. It returns graph structure, not values.

## `trace_dependents`

Input: `path`, `ref_or_symbol`, `depth: int = 2` capped at 8, and
`max_nodes: int = 200`.

Output: a bounded tree of downstream symbols or references and edge reasons,
including a truncation flag. It returns graph structure, not values.

## `trace_path`

Input: `path`, `from_ref_or_symbol`, `to_ref_or_symbol`,
`max_paths: int = 3`, and `max_depth: int = 12`.

Output: bounded block-level dependency paths, each expressed as
`{symbol, via}` steps. No connection returns
`{"connected": false, "paths": []}` rather than an error.

## `explain_formula`

Input: `path`, `ref`.

Output: A1 and R1C1 forms, formula-block ID and extent, resolved names and
structured references, volatile/opaque flags, and diagnostics attached to that
cell.

## `get_diagnostics`

Input: `path`, optional `sheet`, `severity`, and `code`, plus
`max: int = 100`.

Output: matching diagnostics and counts. Workbook diagnostics are distinct from
canonical tool errors such as `E_NOT_FOUND`, `E_INVALID_REF`, and
`E_STALE_CURSOR`.

## `profile_column`

Input: `path`, `col_symbol_or_ref`.

Output: count, non-null, sum, mean, minimum, maximum, and distinct estimate for
numeric columns; otherwise the top five bounded values. A workbook without
cached formula values returns `cachedValues: false` with a recalculation hint.

## `write_cells`

Input: `path` and at most 500 `{ref, value?, formula?}` items. Supported values
are number, string, boolean, and null; datetime values yield
`E_INVALID_VALUE`.

Output: per-cell success/error results and `staleBlocks`. P6 must implement the
surgical OOXML writer, conflict and lock checks, array-formula refusal, and
untouched-part fidelity before P7 exposes this destructive tool.

## `set_column_formula`

Input: `path`, `col_symbol`, `formula`, and `overwrite: bool = false`. Formula
input may be A1 at the anchor or R1C1.

Output: formula-block ID, cells written, and `staleBlocks`. P6 must implement
formula translation, overwrite protection, surgical persistence, and direct
index patching before P7 exposes this destructive tool.

## P7 completion evidence

P7 must add exact generated schemas and worked examples here, then link
`tests/mcp/test_conformance.py` and `docs/evidence/p7-mcp-cli.md`. That evidence
must cover all 14 happy paths, at least one error path per tool, annotations,
initialization instructions, progress, response caps, path confinement,
freshness, pagination, and cursor invalidation.
