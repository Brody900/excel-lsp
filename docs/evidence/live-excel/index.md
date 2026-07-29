# P8 live desktop-Excel evidence

Status: verified P8 evidence; the combined formal reviewer approved both gates
on the exact final frozen fingerprint.

Environment: Windows 11, desktop Microsoft Excel 16.0 build 19530, 64-bit,
with assertions through pywin32 COM and visible user-interface actions captured
from the same Excel session. The protocol implementation is
`tests/live/p8_protocol.py`; the demo builder is
`tests/live/build_p8_demo.py`.

No Trust Center setting was changed. For the local unsigned F16 fixture, the
normal per-workbook **Enable Content** button was used. The workbook and all
screenshots contain only synthetic fixture data.

## Numbered screenshots

| Step | Evidence | What it proves |
|---|---|---|
| 1–2 | [01-index-map.png](01-index-map.png) | The Excel-authored L2 workbook is open beside the exact product map: native table `A1:D4`, confidence `1.0`, four typed columns, totals row included. |
| 2 | [02-diagnostics.png](02-diagnostics.png) | The Excel-authored L3 workbook is open beside structured diagnostics: five `E_ERRVAL` findings for `#REF!` and one `W_INCONSISTENT_FORMULA` finding. Cell C4 visibly contains the hand-edited `=A4*99`. |
| 3 | [03-no-repair.png](03-no-repair.png) | Product-edited L1 opened normally, with no repair dialog, and recalculated to B2:B4 = 60, 80, 120 and C2:C4 = 75, 100, 150. |
| 3 | [04-vba-stamp.png](04-vba-stamp.png) | Product-edited F16 opened normally; the preserved `Stamp` macro ran and visibly wrote `42` to `MacroModel!Z1`. |
| 5 | [05-write-refusal.png](05-write-refusal.png) | L1 remains open beside the structured `E_OPEN_IN_EXCEL` refusal for a write to `Inputs!B3`. |
| 6 | [06-chart-intact.png](06-chart-intact.png) | Product-edited F21 opened without repair; its chart remains rendered and the embedded `Image 2` is selected with visible handles. |
| 4 | [07-trace-a2.png](07-trace-a2.png) | Excel Trace Precedents for `Model!A2` shows the cross-sheet input. |
| 4 | [08-trace-b2.png](08-trace-b2.png) | Excel Trace Precedents for `Model!B2` shows the direct arrow from A2. |
| 4 | [09-trace-c2.png](09-trace-c2.png) | Excel Trace Precedents for `Model!C2` shows direct arrows from A2 and B2. |

## Lineage demo

The short README asset [lineage-demo.gif](../../assets/lineage-demo.gif) is
assembled only from screenshots 07–09. [demo-capture.json](demo-capture.json)
records every source hash, the output hash, dimensions, and frame duration.

## Machine-readable assertions

- [authoring.json](authoring.json) records the native workbook structures and
  Excel version/build. L1 has two sheets, fill-down formulas, and the
  `PrimaryInput` name; L2 has native `Table1` with a totals row; L3 has five
  broken references and one hand-edited filled formula.
- [01-l1-map.json](01-l1-map.json) and
  [01-l2-map.json](01-l2-map.json) are fresh product maps of the Excel-authored
  workbooks.
- [02-l3-diagnostics.json](02-l3-diagnostics.json) is the complete structured
  diagnostic response. Its exact code counts are `E_ERRVAL: 5` and
  `W_INCONSISTENT_FORMULA: 1`.
- [product-writes.json](product-writes.json) records successful
  `write_cells` and `set_column_formula` operations on L1 and F16 and the F21
  cell edit.
- [roundtrip.json](roundtrip.json) records normal opens, no repair dialogs,
  recalculated values, the successful macro assertion, post-save refreshes,
  cleared staleness, the write refusal, and all three product traces.
- [write-refusal.json](write-refusal.json) isolates the exact structured error
  shown in screenshot 05.
- [chart-preservation.json](chart-preservation.json) records one chart object
  and the preserved `Chart 1` and `Image 2` shapes, with actual COM visibility,
  type, position, positive dimensions, and selected-image assertions.

## Trace cross-validation

The COM checks first assert the literal formulas for the chosen row. Excel's
visible arrows then establish the direct row-level precedents. Excel LSP
returns the corresponding three-row formula-block ranges, which preserve the
same relationships across the complete fill-down block:

| Cell | Excel direct precedent(s) | Excel LSP trace child range(s) |
|---|---|---|
| `Model!A2` | `Inputs!B2` | `Inputs!B2:B4` |
| `Model!B2` | `Model!A2` | `Model!A2:A4` |
| `Model!C2` | `Model!A2`, `Model!B2` | `Model!A2:A4`, `Model!B2:B4` |

The script asserts the complete returned sets, tree roots, recalculated values,
and literal formulas. It does not infer successful tracing from screenshots.

## F16 live-discovered fixture repair

The first generated F16 passed byte-preservation tests but Excel displayed an
empty Macro dialog. A real-source comparison reproduced two missing host
requirements that the earlier suffix-only test did not protect:

1. `workbookPr@codeName="ThisWorkbook"` and
   `sheetPr@codeName="Sheet1"`;
2. the exact VBA relationship type
   `http://schemas.microsoft.com/office/2006/relationships/vbaProject`.

The generator and deterministic hash test now freeze both requirements. After
regeneration and the same Excel LSP edits, Excel displayed its normal macro
security banner, listed `Stamp`, ran it, and COM read `MacroModel!Z1 = 42`.
The source workbook remained local-only and was not modified or committed.

## Fresh commands

```text
uv run python tests/live/p8_protocol.py assert-authored
uv run python tests/live/p8_protocol.py inspect-product
uv run python tests/live/p8_protocol.py prepare-writes
uv run python tests/live/p8_protocol.py open-roundtrip
uv run python tests/live/p8_protocol.py open-chart
uv run python tests/live/build_p8_demo.py
```

Each command completed successfully in the final run. The open modes require
the visible desktop-Excel session and are intentionally excluded from CI.
