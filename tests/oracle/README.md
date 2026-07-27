# Parser oracle tests

`oracle.py` independently dual-loads each generated workbook with openpyxl in
read-only mode: once for formulas and once for cached values. It emits canonical
`(sheet, ref, normalized_value, formula)` tuples and compares them with the
production `OOXMLParser` stream. The comparison covers every fixture currently
emitted by `generate_all` (F01, F02, F03, F07, F12, F13, F14, F19, and F20),
including F03's injected cross-sheet formula caches, F13's style-driven dates,
and F19's stored modern-function syntax, `_xlpm.` lexical locals, and cached
results.

Run the pinned read-only behavior probe with:

```text
uv run python -m tests.oracle.probe_openpyxl
```

The probe generates F07 in a temporary directory when no path is supplied, so
it leaves no workbook behind. Its verified limitations and the deliberately
empty cell-stream exception list are recorded in `skiplist.md`.
