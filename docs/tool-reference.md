# Tool reference

Excel LSP exposes exactly 14 local stdio MCP tools. The canonical generated input
schemas and annotations are committed in
[`p7-tool-schemas.json`](evidence/p7-tool-schemas.json); the stdio conformance
test compares that artifact to `tools/list` so schema drift fails the suite.

The first 12 tools are reads with `readOnlyHint: true` and
`openWorldHint: false`. The two writes declare `readOnlyHint: false` and
`destructiveHint: true`. Every call checks freshness. Successful responses are
at most 8,000 serialized characters, no page contains more than 200 raw values,
and expected failures use the canonical `{"error":{"code":"E_...",...}}`
envelope.

The examples use `model.xlsx` as a compact stand-in. Cursors are abbreviated
only in this prose; the server returns the complete opaque base64 token.

## `open_workbook`

Schema key: `open_workbook`. Opens or refreshes an index, emits per-sheet
progress when the client supplies a progress token, and returns the compact map.

```json
{"request":{"path":"model.xlsx"},"response":{"workbook":"model.xlsx","sheets":3,"stale":false,"sheetList":[{"sheet":"Inputs","dims":"A1:C5","regions":[{"id":"region:Inputs:0","range":"A1:C5"}]}],"reindexed":true}}
```

## `refresh`

Schema key: `refresh`. Rechecks the workbook and optionally clears staleness
after Excel recalculation.

```json
{"request":{"path":"model.xlsx","recalculated":true},"response":{"workbook":"model.xlsx","sheets":3,"stale":false,"reindexed":true,"reindexedSheets":["Calc","Summary"]}}
```

## `list_symbols`

Schema key: `list_symbols`. Kinds are `sheets`, `regions`, `columns`,
`names`, `fblocks`, and `cells`. Filtering covers the complete indexed symbol
domain; the response returns the first 100 stable IDs in deterministic order
with a truthful full-match `total` and `truncated` flag.

```json
{"request":{"path":"model.xlsx","query":"revenue","kinds":["columns"]},"response":{"symbols":[{"id":"col:Calc:0:revenue","kind":"column","description":"Revenue (float)"}],"total":1,"truncated":false,"reindexed":false}}
```

## `get_region_schema`

Schema key: `get_region_schema`. Samples are capped by
`max(0, min(3, 180 // ncols))`; a zero-row sample points callers to
`read_range`.

```json
{"request":{"path":"model.xlsx","region_id":"region:Inputs:0"},"response":{"region":"region:Inputs:0","sheet":"Inputs","range":"A1:C5","columns":[{"id":"col:Inputs:0:value","header":"Value","dtype":"float","nonnull":4,"distinct":4,"validation":[]}],"samples":[["Tax rate",0.21,"%"]],"formulaBlocks":[],"conf":1.0,"stale":false,"reindexed":false}}
```

## `read_range`

Schema key: `read_range`. This is the only bulk-value tool. Pages contain at
most 200 cells. The cursor binds the normalized parameters and index generation;
a refresh or write makes it return `E_STALE_CURSOR`.

```json
{"request":{"path":"model.xlsx","ref":"Inputs!A1:C5","max_cells":3},"response":{"sheet":"Inputs","range":"A1:C5","page":{"start":"A1","end":"C1"},"values":[["Input","Value","Unit"]],"offset":0,"totalCells":15,"truncated":true,"cursor":"eyJ...","stale":false,"reindexed":false}}
```

## `find`

Schema key: `find`. Search subjects are limited to 1,000 characters, snippets
to 80, patterns to 256, and matching to two seconds. A timeout returns partial
results plus `W_REGEX_TIMEOUT`. Whenever `values` participates, `stale`
conservatively reports whether any searched cached-value range is invalidated.

```json
{"request":{"path":"model.xlsx","pattern":"Revenue","in":["headers"],"max":10},"response":{"matches":[{"ref":"col:Calc:0:revenue","kind":"header","snippet":"Revenue"}],"truncated":false,"reindexed":false}}
```

## `trace_precedents`

Schema key: `trace_precedents`. Depth is capped at 8 and nodes at 200.

```json
{"request":{"path":"model.xlsx","ref_or_symbol":"Calc!C2","depth":2,"max_nodes":20},"response":{"direction":"precedents","tree":{"kind":"cell","symbol":"cell:Calc!C2","ref":"Calc!C2","via":null,"childCount":1,"children":[{"kind":"cell","symbol":"cell:Inputs!B2","ref":"Inputs!B2","via":"ref","childCount":0,"children":[]}]},"nodeCount":2,"edgeCount":1,"truncated":false,"reindexed":false}}
```

## `trace_dependents`

Schema key: `trace_dependents`. It returns the bounded downstream tree without
cell values.

```json
{"request":{"path":"model.xlsx","ref_or_symbol":"Inputs!B2","depth":1},"response":{"direction":"dependents","tree":{"kind":"cell","symbol":"cell:Inputs!B2","ref":"Inputs!B2","via":null,"childCount":1,"children":[{"kind":"fblock","symbol":"fblock:Calc:3","ref":"Calc!D2:D6","via":"ref","childCount":0,"children":[]}]},"nodeCount":2,"edgeCount":1,"truncated":false,"reindexed":false}}
```

## `trace_path`

Schema key: `trace_path`. It returns bounded shortest block-level paths, or
`{"connected":false,"paths":[]}` when no route exists.

```json
{"request":{"path":"model.xlsx","from_ref_or_symbol":"Inputs!B2","to_ref_or_symbol":"Summary!C2","max_paths":3,"max_depth":12},"response":{"connected":true,"paths":[[{"symbol":"cell:Inputs!B2","via":null},{"symbol":"fblock:Calc:3","via":"ref"},{"symbol":"fblock:Summary:0","via":"ref"}]],"truncated":false,"reindexed":false}}
```

## `explain_formula`

Schema key: `explain_formula`. It reports A1/R1C1 forms, block geometry,
resolution classes, flags, and cell diagnostics.

```json
{"request":{"path":"model.xlsx","ref":"Calc!C2"},"response":{"ref":"cell:Calc!C2","a1":"=B2*Inputs!$B$2","r1c1":"=RC[-1]*Inputs!R2C2","block":"fblock:Calc:2","extent":"C2:C6","resolvedNames":[],"structuredRefs":[],"volatile":false,"opaque":false,"diagnostics":[],"reindexed":false}}
```

## `get_diagnostics`

Schema key: `get_diagnostics`. Filters are applied before the 100-item cap;
counts describe the full filtered set.

```json
{"request":{"path":"model.xlsx","severity":"warn","max":100},"response":{"diagnostics":[{"severity":"warn","code":"W_UNKNOWN_NAME","sheet":"Calc","ref":"cell:Calc!D2","message":"Formula references an unknown name.","related":{}}],"total":1,"counts":{"severity":{"error":0,"warn":1,"info":0},"code":{"W_UNKNOWN_NAME":1}},"truncated":false,"reindexed":false}}
```

## `profile_column`

Schema key: `profile_column`. Numeric columns return aggregates; categorical
columns return the five most common bounded values. Missing formula caches set
`cachedValues: false` and include a recalculation hint. `count` is the number
of positions in the resolved range, including blanks; `nonnull` is the number
of normalized non-null values.

```json
{"request":{"path":"model.xlsx","col_symbol_or_ref":"Inputs!B:B"},"response":{"column":"Inputs!B:B","sheet":"Inputs","range":"B2:B5","count":4,"nonnull":4,"distinct":4,"cachedValues":true,"stale":false,"sum":1000.95,"mean":250.2375,"min":0.1,"max":1000.0,"reindexed":false}}
```

## `write_cells`

Schema key: `write_cells`. Accepts 1–500 qualified edits. Each item supplies
exactly one of `value` or `formula`. Values are number, string, boolean, or
null; formulas begin with `=`.

```json
{"request":{"path":"model.xlsx","cells":[{"ref":"Inputs!B2","value":0.25},{"ref":"Inputs!B3","formula":"=B2*2"}]},"response":{"results":[{"ref":"Inputs!B2","ok":true},{"ref":"Inputs!B3","ok":true}],"staleBlocks":5,"generation":4,"reindexed":false}}
```

## `set_column_formula`

Schema key: `set_column_formula`. The formula may be A1 at the column anchor or
R1C1. Existing cells require explicit `overwrite: true`.

```json
{"request":{"path":"model.xlsx","col_symbol":"col:Calc:0:cost","formula":"=RC[-1]*0.5","overwrite":true},"response":{"formulaBlock":"fblock:Calc:2","cellsWritten":5,"staleBlocks":6,"generation":5,"reindexed":false}}
```

## Canonical errors and executable evidence

Canonical codes are `E_NOT_FOUND`, `E_UNSUPPORTED_FORMAT`, `E_ENCRYPTED`,
`E_LOCKED`, `E_OPEN_IN_EXCEL`, `E_CONFLICT`, `E_CORRUPT`,
`E_STALE_CURSOR`, `E_INVALID_REF`, `E_UNKNOWN_SYMBOL`,
`E_ARRAY_FORMULA`, `E_INVALID_VALUE`, `E_PATH_DENIED`, and
`E_INTERNAL`. The 8,000-character cap applies to canonical error envelopes as
well as successful responses; oversized echoed inputs and details are
deterministically truncated without changing the error code.

[`tests/mcp/test_conformance.py`](../tests/mcp/test_conformance.py) initializes
the real subprocess server through the official MCP client SDK, compares the
generated schemas, calls all 14 happy paths and every tool's error path, verifies
progress, annotations, successful and oversized-error 8k caps, the 200-value
cap, complete late-cell symbol search, cached-value staleness, sparse profile
counts, pagination, write invalidation, and path confinement. The recorded phase results are in
[`p7-mcp-cli.md`](evidence/p7-mcp-cli.md).
