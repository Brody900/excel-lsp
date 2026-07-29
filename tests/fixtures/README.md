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
- Deterministic injection freezes the VBA host code names `ThisWorkbook` and
  `Sheet1` plus Excel's exact relationship type
  `http://schemas.microsoft.com/office/2006/relationships/vbaProject`. The P8
  desktop pass discovered and regression-tested these host requirements after
  a suffix-only metadata assertion allowed Excel to ignore the project.

The local source workbook is intentionally ignored by Git. It is not needed to
regenerate F16; the committed project blob is the deterministic input.

## Implemented generated fixtures

`generate.py` currently emits the implemented P1-P6 subset into `generated/`:

- F01 `basic_single_table.xlsx`: a clean native table with arithmetic formula
  cells and generator-computed numeric caches.
- F02 `multi_region.xlsx`: three islands on one sheet. The first island has one
  intentionally blank body row so `gap_tol=1` joins it and `gap_tol=0` splits
  it; the other islands are separated by wider gaps.
- F03 `cross_sheet_model.xlsx`: the compact reference model, with native tables
  on Inputs, Calc, and Summary and fully injected caches for the
  Inputs-to-Calc-to-Summary chains ending at `Summary!C10`.
- F04 `named_ranges.xlsx`: two global range names plus one Calc-scoped range
  name, with formulas that consume all three names and deterministic caches.
- F05 `structured_table.xlsx`: native `Table1`, four `[@Qty]*[@Price]`
  current-row formulas, explicit `Table1[Col]` consumers, and a totals row
  whose table metadata and formula caches are locked by exact assertions.
- F06 `perf_50k.xlsx` plus `perf_1k.xlsx` and `perf_10k.xlsx`: a ten-column
  streaming-index timing family with one copied/shared formula column and
  generator-computed caches. `generate_all` returns the 50k S1 target under
  F06 and authors all three siblings for the P8 timing series.
- F07 `formula_blocks.xlsx`: two genuine shared-formula groups around a single
  explicit formula tamper, plus caches, a native table, and the harmless
  `E1:F1` merge used by the openpyxl read-only probe.
- F08 `errors.xlsx`: every error text named by the P5 contract plus an
  unrecognized `#FIELD!` cache, all injected as formula results with OOXML
  `t="e"` so detection is proven type-based rather than whitelist-based.
- F09a `circular.xlsx`: a minimal two-cell true cycle with saved zero caches.
- F09b `running_total.xlsx`: a zero seed in `B2` followed by 50,000 formulas in
  `B3:B50002`, each referencing the expanding strictly-earlier range
  `$B$2:B<previous-row>` and carrying a saved zero cache. The block's coarse
  destination overlaps its source even though no concrete cell depends on
  itself, making it the large false-circular and bounded-verification guard.
- F10 `external_link.xlsx`: a genuine numeric `[1]` external-link mapping to a
  missing relative workbook, including the workbook relationship, link part,
  external target relationship, and content-type override.
- F11 `indirect_offset.xlsx`: `INDIRECT` and `OFFSET` reference formulas with
  deterministic caches, exercising both dynamic and volatile diagnostics.
- F12 `merged_headers.xlsx`: a heuristic region with two header rows and three
  merged parent headings.
- F13 `mixed_types.xlsx`: a native table containing integers, numFmt-derived
  dates authored as raw serials, currency, percentages, numbers stored as
  text, booleans, and a deliberately mixed column.
- F14 `sparse.xlsx`: two empty sheets surrounding a sheet with two distant
  singleton cells. Its bounds stay moderate so the independent openpyxl oracle
  does not turn a sparse test into a dense stress test.
- F15 `threeD_ref.xlsx`: Jan, Feb, and Mar source sheets plus the exact
  `SUM(Jan:Mar!B2)` 3-D consumer and its saved result on Summary.
- F16 `macro_book.xlsm`: a macro-enabled worksheet with the sanctioned
  `vbaProject.bin`, one cached arithmetic formula, and the live `Stamp` target;
  P6 edits prove the VBA part is byte-identical afterward.
- F17 `unicode_names.xlsx`: non-ASCII values and worksheet names, including an
  apostrophe that requires doubled quoting in formulas, plus a Unicode defined
  name and deterministic cross-sheet formula caches.
- F18 `volatile.xlsx`: independent `NOW` and `RAND` formula blocks with stable
  saved caches, proving one `I_VOLATILE` finding per block.
- F19 `modern_functions.xlsx`: stored `_xlfn.`/`_xlws.` functions and Excel's
  `_xlpm.` lexical-local spelling, multiple LET bindings, a LAMBDA defined name
  invoked as a function, XLOOKUP, literal `A1#`
  and defined-name spill consumers, `@` implicit intersection, a saved spill
  follower value, and generator-computed caches for every formula.
- F20 `stress_map.xlsx`: 40 sheets, 12 distinctly sized regions on the first
  sheet, hidden and very-hidden sheets, and 300 global defined names split
  evenly across range, multi-range, constant, formula, and lambda kinds. Its
  generated LAMBDA names also preserve `_xlpm.` parameter spelling.
- F21 `chart_image.xlsx`: a native bar chart and embedded deterministic PNG on
  a small dashboard sheet; P6 part diffs freeze the chart, drawing,
  relationship, and media parts across edits.

Every OOXML member is repacked in lexical order with ZIP timestamp
`1980-01-01 00:00:00`. Workbook document properties use `2000-01-01 00:00:00`.
The generator never evaluates arbitrary formulas; it injects only results it
computes directly from the values it authored. Tests lock the SHA-256 values of
F01, F07, F19, every implemented P4 fixture, and all four P5 fixtures, and
independently assert that two complete corpus generations are byte-identical.
Shape tests and oracle tests share one generated corpus per module so F09b is
not needlessly authored for every individual assertion; the byte-determinism
test still performs two independent complete generations.

`render_part_diff.py F16|F21` applies the frozen P6 edit to a temporary fixture
copy and prints the complete before/after part manifest used by
`docs/evidence/part-diff-f16.json` and `part-diff-f21.json`. Unit tests require
fresh renderer output to equal both committed artifacts exactly.
