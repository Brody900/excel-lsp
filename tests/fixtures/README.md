# Excel LSP fixture provenance

Most fixture workbooks are generated deterministically by `generate.py` and are
not committed. The sole sanctioned binary source asset is
`assets/vbaProject.bin`, which is injected into generated fixture F16 so the
surgical writer can prove that it preserves VBA content byte-for-byte.

## `vbaProject.bin`

- Authored in desktop Microsoft Excel 16.0 (build 19530, 64-bit) on Windows 11.
- Source workbook: local-only `f16_source.xlsm`, created on 2026-07-15.
- The VBA project contains exactly this test macro:

  ```vb
  Sub Stamp()
      Range("Z1").Value = 42
  End Sub
  ```

- Extracted OOXML member: `xl/vbaProject.bin`.
- Source workbook SHA-256:
  `cf1c3016d5409905a33fbdbceaf2612f060dbba13274438a0936fbcca1aadb46`.
- Extracted blob SHA-256:
  `be05aafbb31d2de0ffd686c9cae71b97a2596132ed4443d9d656558d7089ccb1`.
- Macro source was independently inspected with `olevba 0.60.2` before the
  asset was accepted. The live Excel release gate also executes `Stamp` and
  asserts that cell `Z1` becomes `42` after Excel LSP edits F16.

The local source workbook is intentionally ignored by Git. It is not needed to
regenerate F16; the committed project blob is the deterministic input.

## Implemented generated fixtures

`generate.py` currently emits the P1/P2 subset into `generated/`:

- F01 `basic_single_table.xlsx`: a clean native table with arithmetic formula
  cells and generator-computed numeric caches.
- F02 `multi_region.xlsx`: three islands on one sheet. The first island has one
  intentionally blank body row so `gap_tol=1` joins it and `gap_tol=0` splits
  it; the other islands are separated by wider gaps.
- F03 `cross_sheet_model.xlsx`: the compact reference model, with native tables
  on Inputs, Calc, and Summary and fully injected caches for the
  Inputs-to-Calc-to-Summary chains ending at `Summary!C10`.
- F07 `formula_blocks.xlsx`: two genuine shared-formula groups around a single
  explicit formula tamper, plus caches, a native table, and the harmless
  `E1:F1` merge used by the openpyxl read-only probe.
- F12 `merged_headers.xlsx`: a heuristic region with two header rows and three
  merged parent headings.
- F13 `mixed_types.xlsx`: a native table containing integers, numFmt-derived
  dates authored as raw serials, currency, percentages, numbers stored as
  text, booleans, and a deliberately mixed column.
- F14 `sparse.xlsx`: two empty sheets surrounding a sheet with two distant
  singleton cells. Its bounds stay moderate so the independent openpyxl oracle
  does not turn a sparse test into a dense stress test.
- F20 `stress_map.xlsx`: 40 sheets, 12 distinctly sized regions on the first
  sheet, hidden and very-hidden sheets, and 300 global defined names split
  evenly across range, multi-range, constant, formula, and lambda kinds.

Every OOXML member is repacked in lexical order with ZIP timestamp
`1980-01-01 00:00:00`. Workbook document properties use `2000-01-01 00:00:00`.
The generator never evaluates arbitrary formulas; it injects only results it
computes directly from the values it authored. Extending the corpus preserves
the P1 fixture archives byte-for-byte; tests lock their SHA-256 values as well
as checking that two complete generations are identical.
