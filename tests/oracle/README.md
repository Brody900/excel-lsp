# Parser oracle tests

`oracle.py` independently dual-loads each generated workbook with openpyxl in
read-only mode: once for formulas and once for cached values. It emits canonical
`(sheet, ref, normalized_value, formula)` tuples and compares them with the
production `OOXMLParser` stream. The comparison covers every fixture currently
emitted by `generate_all` (F01, F02, F03, F04, F05, F07, F09a, F09b, F12,
F13, F14, F15, F19, and F20), including F03's cross-sheet caches; F04's
scoped/global-name consumers; F05's structured and totals-row formulas; all
50,000 F09b running-total formula caches; F13's style-driven dates; F15's 3-D
formula; and F19's stored modern-function syntax, `_xlpm.` lexical locals, and
cached results.

The oracle module authors the complete corpus once and reuses it across tests.
This keeps the 50,000-row F09b contract fully covered without regenerating the
large workbook for each separate oracle assertion.

Run the pinned read-only behavior probe with:

```text
uv run python -m tests.oracle.probe_openpyxl
```

The probe generates F07 in a temporary directory when no path is supplied, so
it leaves no workbook behind. Its verified limitations and the deliberately
empty cell-stream exception list are recorded in `skiplist.md`.
