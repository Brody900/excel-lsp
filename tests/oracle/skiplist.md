# T-oracle skip list

Pinned probe runtime: openpyxl 3.1.5. Reproduce with:

```text
uv run python -m tests.oracle.probe_openpyxl
```

## Active cell-stream skips

None. In particular, openpyxl 3.1.5 expands both F07 shared-formula groups in
read-only mode: `C3`, `C11`, `C14`, and `C21` return their correctly translated
formula text. Shared followers therefore require no oracle exception.

## Verified read-only metadata deficiencies

These observations narrow the cell-stream oracle's scope; they are not silent
cell exclusions:

1. `ReadOnlyWorksheet.tables` raises `AttributeError`. The normal-mode control
   sees `FormulaBlocksTable`, proving that the fixture contains the ListObject.
2. `ReadOnlyWorksheet.merged_cells` raises `AttributeError`. The normal-mode
   control sees `E1:F1`, proving that the fixture contains the merged range.

The production lxml parser must parse table parts and merged ranges directly.
No unverified or anticipated deficiency belongs in this file.
