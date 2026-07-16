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
