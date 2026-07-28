# P6 surgical editor and staleness evidence

P6 implements the only production workbook-mutation boundary in Excel LSP.
It edits OOXML parts directly, patches the semantic index without a whole-sheet
reindex, and records transitive dependent blocks as stale until Excel saves a
recalculated workbook. The exact-fingerprint combined formal gate approved the
fully remediated candidate below.

## Delivered contracts

- `write_cells` accepts at most 500 typed value or formula edits. Values are
  finite numbers, strings, booleans, or null; datetime values are rejected.
- `set_column_formula` resolves one frozen `col:` symbol, requires explicit
  overwrite for occupied bodies, and fills an A1-anchor pattern or R1C1
  pattern across the complete body. It is deliberately not constrained by the
  separate 500-cell `write_cells` batch limit.
- Existing workbooks are never loaded and saved through openpyxl in production.
  The writer copies ZIP entries and changes only the declared worksheet and
  calculation parts.
- A sibling Excel lockfile is checked before the source-hash conflict check.
  Replacement uses a same-directory temporary archive, file `fsync`, a parser
  validation pass, a second source-hash check, and atomic replacement with one
  permission retry. Before each replacement attempt, including the delayed
  retry, lockfile precedence and the indexed destination hash are revalidated.
- The sparse patch verifies that the workbook still has the writer's exact
  installed hash. Stable path hash/stat snapshots bracket touched-sheet
  collection, the direct-patch transaction, and its commit so parser content,
  package hashes, and source metadata belong to one file generation. A forced
  post-replacement second-writer race rebuilds the sidecar from current bytes
  and returns `E_CONFLICT` instead of reporting a potentially overwritten edit
  as successful.
- Every successful edit updates sparse cell rows, region summaries, formula
  blocks, graph mirrors, part hashes, source stat metadata, and generation in
  one index transaction. A post-replacement index failure triggers an
  incremental recovery and preserves the stale plan rather than claiming a
  direct patch succeeded.

## Cell and calculation mechanics

Worksheet rows and cells are indexed and required to be ordered. New elements
are inserted at their sorted position. Existing style attributes are retained;
new cells have no style. Numbers use an untyped `<v>` and integral floats omit
the decimal suffix. Booleans use `t="b"`, strings use an inline string with
`xml:space="preserve"`, null leaves an empty cell, and formula writes remove
the stale cache. Text limits use UTF-16 code units, and invalid Unicode or XML
characters fail as `E_INVALID_VALUE` before source replacement.
Replacement `<f>`, `<v>`, and `<is>` children are inserted before a preserved
cell-level `<extLst>`, retaining the `CT_Cell` schema order.

Editing any shared-formula member expands the complete declared group to
ordinary translated formulas before applying the requested edit. An incomplete
or non-unique group is corrupt; any target inside a multi-cell array span is
refused with `E_ARRAY_FORMULA`. A write extending actual cells recomputes
`<dimension>`; a within-extent write drops that advisory element.

The writer ensures workbook `calcPr` has `fullCalcOnLoad="1"`. When `calcPr`
is absent, it is inserted before every valid later workbook child (`oleSize`
through `extLst`) rather than appended blindly. If a calculation chain exists,
its part, workbook relationship, and content-type override are removed
together. `xl/sharedStrings.xml` is never mutated.

Primary tests:

- `tests/unit/test_editor.py::test_surgical_write_changes_only_declared_parts_and_serializes_cell_types`
- `tests/unit/test_editor.py::test_editing_shared_follower_expands_the_complete_group_before_write`
- `tests/unit/test_editor.py::test_multi_cell_array_formula_refuses_anchor_and_follower_without_writing`
- `tests/unit/test_editor.py::test_calc_chain_part_relationship_and_override_are_deleted_together`
- `tests/unit/test_editor.py::test_lockfile_precedes_conflict_and_neither_failure_mutates_workbook`
- `tests/unit/test_editor.py::test_excel_text_and_numeric_boundaries_reject_without_mutation`
- `tests/unit/test_editor.py::test_cell_payloads_remain_before_preserved_extension_lists`
- `tests/unit/test_editor.py::test_new_calc_properties_precede_every_later_workbook_child`
- `tests/unit/test_editor.py::test_replace_retry_preserves_intervening_external_workbook`
- `tests/unit/test_editor.py::test_replace_retry_rechecks_excel_lock_before_destination_hash`
- `tests/unit/test_editor.py::test_replace_retry_succeeds_when_preconditions_remain_unchanged`
- `tests/unit/test_editor.py::test_repeated_replace_failure_returns_locked_without_mutation`

## Direct index patch and staleness

Before mutation, `IndexStore.plan_staleness` follows range-overlap dependents
breadth-first at formula-block granularity, with the frozen 50,000-block cap.
The written formula rectangle is also stale because its cached value was
removed. The post-write transaction reparses only touched worksheets, selects
only changed/expanded cells for direct replacement, recomputes derived regions
and formula analysis for those sheets, persists `I_STALE` rectangles, replaces
package hashes, and increments generation once.

An ordinary later refresh clears stale state for each worksheet whose part hash
changed through an external save. `refresh(recalculated=true)` clears all stale
rectangles and `I_STALE` diagnostics even when workbook bytes are unchanged,
bumping generation only when that clear mutated the index.

Primary tests:

- `tests/unit/test_editor.py::test_write_service_directly_patches_index_and_marks_transitive_dependents_stale`
- `tests/unit/test_editor.py::test_formula_write_removes_cache_and_marks_the_written_formula_stale`
- `tests/unit/test_editor.py::test_post_replace_direct_patch_failure_recovers_index_and_staleness`
- `tests/unit/test_editor.py::test_post_replace_workbook_race_reconciles_index_and_returns_conflict`
- `tests/unit/test_editor.py::test_race_during_sheet_collection_recovers_current_workbook_generation`
- `tests/unit/test_editor.py::test_refresh_recalculated_clears_all_staleness_and_bumps_once`
- `tests/unit/test_editor.py::test_external_recalculation_save_clears_staleness_for_changed_sheet`

## Column-formula translation

The R1C1 renderer supports relative, absolute, and mixed cell coordinates,
ranges, whole rows, whole columns, and qualified references. It rejects mixed
A1/R1C1 input and translations outside worksheet bounds. A1 patterns are
translated from the body anchor. The complete fill is installed in one
generation and returns the newly resolved formula-block symbol. A 501-row
ListObject regression proves this tool does not accidentally inherit the bulk
write limit.

Primary tests:

- `tests/unit/test_from_r1c1.py`
- `tests/unit/test_editor.py::test_set_column_formula_requires_explicit_overwrite_and_fills_a1_pattern`
- `tests/unit/test_editor.py::test_set_column_formula_renders_r1c1_per_body_row_and_rejects_boundary_escape`
- `tests/unit/test_editor.py::test_set_column_formula_is_not_limited_by_write_cells_batch_cap`

## Part preservation

F16 and F21 are deterministic generated fixtures. The committed manifests list
every ZIP part's before/after SHA-256, not only selected protected files. Both
edits change exactly `xl/worksheets/sheet1.xml`; every other member remains
byte-identical. F16 preserves `xl/vbaProject.bin`; F21 preserves its chart,
drawing, drawing relationship, and embedded PNG.

- F16 source SHA-256:
  `9f392ed6227358c5ff79a667433e79290e5acd2b6415e2ea6e80006b09cb6ae6`
- F21 source SHA-256:
  `d8f0d32d51ce7aa864f47c1ad83a192eb650559ee6ac800e297562d3eb0212fb`

- [`part-diff-f16.json`](part-diff-f16.json)
- [`part-diff-f21.json`](part-diff-f21.json)
- `tests/fixtures/render_part_diff.py`
- `tests/unit/test_editor.py::test_committed_part_diff_evidence_matches_fresh_render`

The Hypothesis property executes 50 deterministic random scripts across small
workbooks. It proves every non-deliberate part is byte-identical and reparses
each requested value or formula to the exact normalized semantic result. This
protects I18 without manufacturing a combinatorial matrix unrelated to a
behavioral boundary.

- `tests/property/test_editor_preservation.py::test_i18_random_edit_scripts_preserve_parts_and_reparse_exactly`

## Recalculation boundary

Excel LSP does not implement a formula engine and never invents post-edit
cached values. Formula writes remove `<v>`, and the affected formula and its
transitive dependents remain explicitly stale until an external calculation
save or a caller-confirmed recalculated refresh. The live smoke below proves
that desktop Excel is the compatible calculation boundary for the reference
model.

## Desktop Excel smoke

On 2026-07-28, the live-marked P6 test copied F03, removed workbook `calcPr`,
added a valid later `fileRecoveryPr`, and added a cell-level `extLst` after
`Inputs!B2`. It then changed that cell from `0.10` to `0.20`, asserted the two
repaired schema orders, and proved `Summary!C10` was stale before Excel opened
it. Desktop Excel 16.0 build 19530 opened the adversarial workbook normally,
performed a full calculation, produced the exact expected value `2232.48`,
saved, and closed. A `recalculated=true` refresh then read the saved cache and
cleared staleness.

- [`live-excel/p6-smoke.json`](live-excel/p6-smoke.json)
- `tests/live/test_editor_excel.py::test_surgical_edit_round_trips_through_desktop_excel`
- Fresh command: `uv run pytest tests/live/test_editor_excel.py -m live -q`
- Result: 1 passed in 2.15 seconds with explicit COM proxy release before
  apartment teardown

The P8 protocol still owns Excel-authored L1-L3, the VBA execution assertion,
open-in-Excel refusal evidence, chart screenshot, repair-dialog screenshot, and
demo capture. This P6 smoke does not claim those later gates.

## Candidate verification

Fresh results before the formal freeze:

| Check | Result |
|---|---|
| Full uninstrumented repository suite | 2,035 passed, 1 live deselected in 228.31 seconds |
| Full branch-instrumented repository suite | 2,035 passed, 1 live deselected in 426.04 seconds; 89.65% total core branch coverage |
| Editor/R1C1/property/fixture/oracle slice | 88 passed in 24.74 seconds |
| Editor unit module after final hardening | 50 passed in 4.04 seconds |
| Desktop Excel live schema-order smoke | 1 passed in 2.15 seconds; clean COM teardown |
| Pyright | 0 errors, 0 warnings |
| Ruff lint and format | passed; 86 files formatted |
| Deterministic fixture regeneration | passed; all 20 current fixture IDs emitted with no tracked drift |
| Lock and distribution build | `uv lock --check` resolved 70 packages; sdist and wheel built successfully |
| Git whitespace check | passed |

Each formal R-mech and R-test verdict is accounted only after the single
combined reviewer evaluates its exact frozen fingerprint.

## Formal phase gate

The first combined P6 review evaluated base
`672e91ad362dcf1d984307fe4c2fd9db89d72b04`, staged tree
`fd797eb9796ad561899c8ab9439e44fbd45619a4`, and cached diff
`17eba11d65fbde3e64854170c2cb9e4f4b6ca8e2` unchanged at entry and exit. It
charged R-mech #30 and R-test #25 as `REVISE`: replacement cell content could
follow a preserved `extLst`, and a newly created `calcPr` considered only
`extLst` rather than every valid schema successor. The main agent implemented
both ordered insertions and added exact unit plus desktop-Excel adversarial
coverage. The second review confirmed both schema-order remediations.

That second combined review evaluated base
`672e91ad362dcf1d984307fe4c2fd9db89d72b04`, staged tree
`5671fbba8cf84eed0cf4a1641908e47fd29832dd`, and cached diff
`845e1fd8a8973c718611c802482c0543c980aac9` unchanged at entry and exit. It
charged R-mech #31 and R-test #26 as `REVISE`: after a first replacement
attempt raised `PermissionError`, the delayed retry did not recheck either the
Excel lockfile or the destination hash, and the retry's four material outcomes
had no deterministic regressions. The main-agent remediation revalidates the
lockfile first and then the indexed hash before every attempt, preserving an
intervening writer. Dedicated tests cover external mutation, lockfile
precedence, unchanged retry success, and repeated `E_LOCKED` failure. The next
review confirmed the retry-race remediation and all four regressions.

The third combined review evaluated base
`672e91ad362dcf1d984307fe4c2fd9db89d72b04`, staged tree
`2f0755b91417deccb299a33662d8b11de8663106`, and cached diff
`861802ee5ce0b477df1834409cb8a5f59a779d17` unchanged at entry and exit. It
charged R-mech #32 and R-test #27 as `REVISE`: a second writer could change the
path after the service's installed-hash comparison while the open parser
streamed touched sheets, allowing old parser cells and hashes to be paired with
a later path stat tuple. Stable SHA-256 plus identity/stat snapshots now bracket
collection, direct patch application, and transaction commit. A collection
failure is converted to `E_CONFLICT` when the path generation changed; recovery
then rebuilds from current valid bytes. A deterministic regression installs an
observably changed valid workbook during collection and proves its bytes and
indexed value survive, followed by an ordinary no-change refresh. The next
review below evaluated that frozen candidate.

The fourth combined review evaluated base
`672e91ad362dcf1d984307fe4c2fd9db89d72b04`, staged tree
`8c794e703c0d0a783f23095c89d920d29f3c8f10`, and cached diff
`6587b6161e4c15cde11dfcd10077d1bd323ad266` unchanged at entry and exit. It
charged R-mech #33 and R-test #28 as `APPROVE`, with no findings. The reviewer
independently passed all 50 editor tests and confirmed the sidecar-generation,
atomic-retry, schema-order, preservation, and live-evidence remediations on one
unchanged fingerprint. P6's formal phase gate is closed.
