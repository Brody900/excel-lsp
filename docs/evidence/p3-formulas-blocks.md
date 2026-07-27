# Phase 3 formula references and block evidence

Recorded on 2026-07-16 (America/Los_Angeles) and refreshed on 2026-07-27
(America/New_York). This report covers the P3 candidate for M3 reference
extraction and M4 R1C1 formula blocks. Formal
R-mech and R-test verdicts are recorded only after they inspect the frozen
staged candidate.

## Delivered contracts

- Formula tokenization preserves source spelling, whitespace, quoted sheet
  names, postfix spill operators, implicit intersection, and Excel structured
  header escapes. It contains malformed formulas as one opaque parse result
  rather than aborting a worksheet.
- A1 classification covers local and cross-sheet cells/ranges, quoted names,
  apostrophes, 3-D spans, numeric and direct external-workbook qualifiers,
  workbook- and sheet-scoped names, multi-area names, formula/LAMBDA names,
  and supported structured selectors.
- Public external-link labels pass through one URL-aware sanitizer shared by
  the workbook map and formula edges. Raw targets remain internal lifecycle
  context and are excluded from canonical exports.
- The frozen built-in catalog combines `openpyxl==3.1.5`'s formula set with a
  committed 171-name modern supplement captured from Microsoft's alphabetical
  function catalog on 2026-07-16. Stored `_xlfn.` and `_xlws.` prefixes are
  normalized before function or defined-name lookup. Excel's private `_xlpm.`
  prefix is handled separately: it remains part of the exact lexical lookup
  key, but is removed from public callable labels and duplicate-name identity.
- LET declarations become visible only after their paired value expression;
  prior lexical bindings remain visible inside nested values. LAMBDA parameters
  are lexical bindings. Duplicate or invalid local names are contained as parse
  problems, while Excel's valid missing LET values—including an empty final
  calculation—do not create spurious warnings.
- Raw declarations follow installed worksheet-entry behavior: a letter,
  underscore, or backslash may begin the name; valid in-grid A1 spellings are
  rejected; R1C1-like and beyond-grid A1-looking spellings are accepted. In
  saved OOXML, `_xlpm.` is authoritative and disambiguates even `_xlpm.A1`.
  Unprefixed built-in calls retain built-in precedence over same-spelled raw
  bindings, while explicitly prefixed stored uses resolve only in the private
  local namespace.
- `INDIRECT` and `OFFSET` are always opaque dynamic references. `CHOOSE` and
  `INDEX` are opaque only in syntactic reference contexts. Reference-result
  inference covers parenthesized unions, intersections, nested
  `INDEX`/`CHOOSE`/`OFFSET`/`INDIRECT` sources, absorbed range endpoints, and
  implicit intersections. The complete expression must remain reference-pure,
  so arithmetic, arrays, scalar names, and scalar functions do not promote an
  enclosing call.
- Reference identity is typed through lexical LET values and LAMBDA closures.
  Inline and defined LAMBDAs can be passed as arguments, returned as values,
  stored in later LET bindings, and selected by `CHOOSE` or `IF`; the analysis
  conservatively joins callable alternatives. Named `Apply`, `Pick`, and
  `Make`-style higher-order flows retain whether the final value can be a
  reference, while scalar controls remain non-reference-valued.
- A computed callable result used as either side of `:` is never silently
  accepted as a complete static range. It contributes one visible
  `opaque:<callable>` edge plus `I_DYNAMIC_REF`, including through transparent
  grouping parentheses and named or lexical callable invocation. Simple
  reference-preserving LET/LAMBDA/implicit-intersection wrappers retain a
  nested `INDEX` attribution; deeper computed compositions conservatively
  attribute the outer callable.
- Concrete range names can participate in exact colon/intersection geometry.
  Constants and formula/LAMBDA names are typed separately: their body
  precedents remain visible dependencies but are never mistaken for the value
  returned by the name. When such a name is a composite endpoint, the operator
  stays conservative; `INDEX`, `CHOOSE`, `OFFSET`, `INDIRECT`, LET wrappers,
  and alias chains retain their function-specific `opaque:<FN>` edge and
  `I_DYNAMIC_REF`. Scalar or bare-callable results retain `opaque:ref`.
- Whitespace intersection is a reference context just like colon. Direct,
  grouped, named-LAMBDA, and reversed computed endpoints retain dynamic
  attribution instead of silently degrading to their input precedents. Every
  intersection in one formula is accumulated independently, preserving source
  order and occurrence-specific dynamic diagnostics.
- A lexical local used in `:` or whitespace intersection retains concrete
  precedents plus one conservative `opaque:ref`, including through transparent
  parentheses and in either endpoint order. Scalar locals do not acquire false
  dynamic identity; reference-valued locals still propagate the enclosing
  LET/LAMBDA result attribution.
- Volatile calls mark the block. Dynamic reference calls emit
  `I_DYNAMIC_REF`; unparseable or unknown names emit contained warnings.
  Complete diagnostic aggregation remains P5 work.
- Shared-formula followers use Excel LSP's tokenizer-backed A1 translator
  rather than relying on openpyxl Translator for modern syntax. This covers
  `@`, spill operands, structured escapes, range endpoints absorbed into a
  following function token, and `@`/`#` bound independently to either endpoint
  of one range. The same quote-, bracket-, and 3-D-aware endpoint parser is
  shared by classification, translation, and R1C1 normalization.
- Every formula cell is normalized independently to R1C1. Column-major equal
  runs grow vertically and merge across exactly matching adjacent columns.
  Malformed formulas remain singleton-capable opaque patterns.
- Relative edge rectangles are extruded over the source block, absolute
  coordinates stay fixed, and every result is clamped to Excel's row and
  column bounds. Spill edges deliberately target only the definition anchor;
  v0.1 does not infer dynamic spill extents.
- `W_INCONSISTENT_FORMULA` compares maximal vertical and horizontal formula
  runs of at least five cells. A minority is reported only when the dominant
  pattern is at least 80% and the minority count is at most
  `max(3, ceil(0.05*n))`.

## R1C1 and block evidence

The committed properties establish the P3 block invariants:

| Invariant | Executable evidence |
|---|---|
| I7 formula parsing never aborts indexing | `test_malformed_formulas_return_one_opaque_parse_edge`, recursive-name containment, parser shared-formula corruption tests, and formula-analysis transaction rollback |
| I8 quoted sheet names round-trip | the quoted/apostrophe/3-D classification matrix, qualifier-preservation known answers, and shared-translation qualifier test |
| I9 every formula cell has exactly one block owner | `test_i9_i11_formula_blocks_cover_exactly_once` checks every planted coordinate and rejects overlap |
| I10 translated formulas retain one R1C1 signature | 300 deterministic Hypothesis examples compare against pinned openpyxl Translator; a second 300-example property compares the modern-safe shared translator with openpyxl on its supported grammar |
| I11 block rectangles cover exactly the formula cells | the same non-vacuous random-grid property checks exact owner-set equality, ragged gaps, and deterministic rebuilds |
| I20 the F19 modern matrix has no false parse/name warnings | F19 persists seven exact blocks and eight exact edges while excluding `W_PARSE` and `W_UNKNOWN_NAME` |

F07 is the exact inconsistency oracle. Its two shared fill-down groups plus one
explicit tamper produce three blocks: rows 2–11 multiplication, row 12
addition, and rows 13–21 multiplication. Only `C12` receives
`W_INCONSISTENT_FORMULA`, related to block 0 and the expected multiplication
signature.

Block construction remains exact rather than heuristic. A discarded
optimization tried to reuse one openpyxl Translator across a fill-down run;
preflight found counterexamples for `=@A2`, `=A2:INDEX(A:A,2)`, and the final
Excel row. A custom fused lexer was also discarded after adversarial name
tokens such as `C$a4` diverged from the tokenizer. The accepted implementation
normalizes each cell independently.

The R1C1 signature is a coarse block bucket, not the final merge decision.
Every proposed vertical extension and horizontal merge with noncanonical
spelling is directionally checked by translating the deterministic top-left
formula to the candidate coordinate and requiring exact formula equality. This
prevents lowercase/absolute A1 spelling and nested degenerate-range aliases
from collapsing even though a symmetric canonical key cannot model Excel's
directional case normalization. Canonical uppercase formulas take a proven
fast path. Coordinate spills retain §5.4's explicit exception: identical raw
spill formulas can share one block and one definition-anchor edge, while
different spill-anchor spellings remain separate because spill operands pass
through R1C1 verbatim.

## Composite reference grammar

Classification, shared-formula translation, and R1C1 normalization use one
quote-, bracket-, and 3-D-aware endpoint grammar. It recognizes A1 cells,
rectangles, whole rows and columns, defined names, structured operands, and
independent `@`/`#` endpoint modifiers without splitting quoted or escaped
content.

Static colon operators fold through transparent parentheses and implicit
intersection groups. Compatible endpoint geometries receive their exact Excel
bounding rectangle, including grouped whole-column forms and
A1/concrete-range-name/structured combinations. Both sides may independently
use grouped `@(...)` without
creating an empty callable. When a column-like token is also a defined name,
classification uses workbook context to resolve the ambiguity; context-free
translation and R1C1 conversion preserve the unresolved spelling.

Representable static intersections are reduced to their exact common
rectangle, including parenthesized operands and whole axes. A mixed-coordinate
intersection that cannot be represented by the frozen geometry model retains
its visible endpoint references and adds `opaque:ref`; it never disappears or
pretends to be complete. Multi-area and cross-sheet range operands are likewise
kept explicitly conservative.

The committed matrix covers these boundaries in
`test_static_reference_intersection_is_indexed_as_the_exact_overlap`,
`test_unrepresentable_static_intersection_is_visibly_conservative`,
`test_computed_reference_endpoints_are_conservatively_opaque`,
`test_index_provenance_survives_reference_preserving_wrappers`,
`test_grouped_named_and_lexical_callable_range_endpoints_are_opaque`,
`test_computed_defined_names_remain_conservative_composite_endpoints`,
`test_computed_reference_intersections_emit_dynamic_attribution`,
`test_every_unfolded_intersection_retains_its_own_conservative_attribution`,
`test_nested_computed_name_endpoint_deduplicates_dynamic_issue`,
`test_name_expansion_cache_is_isolated_across_inherited_execution_contexts`,
`test_computed_name_alias_chain_is_bounded_and_keeps_conservative_edges`,
`test_computed_name_endpoints_do_not_fabricate_static_block_hulls`,
`test_computed_name_endpoint_persists_precedents_without_a_false_hull`,
`test_invalid_structured_columns_and_empty_table_sections_are_contained`,
`test_structured_range_endpoints_synthesize_exact_compatible_bounds`,
`test_grouped_whole_column_range_endpoints_keep_exact_bounds`, and
`test_name_and_structured_range_endpoints_share_exact_compatible_bounds`.
Additional higher-order controls are frozen by
`test_higher_order_lambdas_retain_callable_and_reference_results`,
`test_defined_lambda_preserves_callable_arguments`, and
`test_callable_selector_results_retain_conservative_reference_flow`.

## Structured-reference context

A bare structured operand such as `[@Input]` can have different meaning across
two adjacent ListObjects even when every formula has the same R1C1 signature.
P3 preserves one formula block and projects its source-dependent classifications
over homogeneous source rectangles. Each rectangle receives a complete formula
analysis so an `INDEX` or `CHOOSE` that becomes reference-valued only inside a
table still contributes its block-level opaque edge and `I_DYNAMIC_REF`.
Direct-operand provenance keeps ordinary static edges whole and prevents a
structured reference reached through a defined-name body from being
incorrectly reclassified as if it appeared in the worksheet formula.

Two deterministic Hypothesis differentials compare block-level edge coverage
with brute per-formula-cell classification across randomly tiled ListObjects,
including header and totals boundaries. A third 200-example property proves
that the semantic rectangle partition covers every source coordinate exactly
once and never crosses a change in table context, including overlapping
rectangles and duplicate labels.

A separate pre-gate differential compared the resulting block edges and
dynamic/parse diagnostics with the union of per-cell analysis for 720
deterministic randomized table layouts and six contextual `INDEX`/`CHOOSE`
formula families; it found no mismatch. This is engineering preflight evidence,
not a formal review verdict.

The partitioner is a row-event sweep over bounded worksheet columns. It updates
only columns touched by starting or ending contexts and closes adjacent
vertical strips with identical histories. A preflight staggered-table series
measured 128, 256, 512, 1,024, 2,048, and 4,096 contexts in approximately
0.002, 0.005, 0.005, 0.012, 0.051, and 0.055 seconds on the development
machine. The prior active-context rescan took about 5.16 seconds at 1,024.
These figures are engineering smoke evidence, not a release benchmark.

## Spill reference scope

`A1#` and a spilled defined-name consumer resolve to the definition anchor and
use `via="spill"`. Block extrusion is intentionally disabled for spill edges:
the source block may be many cells, but P3 emits only the one statically known
anchor cell. The saved spill follower in F19 is a cached value, not a claimed
statically inferred spill rectangle.

When a range operator has modern syntax on one or both endpoints—such as
`A1#:B5`, `A1:@B5`, or a quoted 3-D equivalent—static `@` endpoints form the
ordinary exact range, while each spilled endpoint contributes only its
definition anchor. Copy-down translation shifts both endpoints. R1C1 keeps a
spill endpoint verbatim as required by §5.4, so formulas with different spill
anchors do not collapse into one false block.

This limitation is executable in
`test_spill_edge_resolves_only_the_anchor_instead_of_extruding_the_block` and
the exact F19 persistence test.

## Persistence and lifecycle

Schema version 3 adds normalized internal `list_objects` and
`list_object_columns` tables so structured-reference analysis retains exact
header, totals, and column context after the parser stream closes. Formula
analysis atomically replaces source-owned blocks, edges, and P3 diagnostics,
then links region columns to blocks only when one block covers the complete
ListObject data body.

Table `name` and `displayName` aliases share one case-insensitive,
workbook-wide namespace. Fresh and incremental indexing reject collisions as
structured `E_CORRUPT` errors and roll back the complete canonical export and
generation; persisted catalogs are defensively revalidated before formula
analysis.

The lifecycle distinguishes changed formula sources from changed semantic
context:

- a changed worksheet reanalyzes that formula source while preserving incoming
  edges from unchanged sheets;
- a changed defined-name or ListObject catalog reanalyzes every formula source
  whose meaning may depend on it;
- an external-link relationship target can change with no worksheet reparse,
  while all formula sources are reclassified against the new sanitized label;
- an equal incremental result and a fresh full index have identical
  natural-key canonical exports;
- an injected second-sheet analysis failure rolls back blocks, edges,
  diagnostics, metadata, and generation together.

Before a multi-sheet refresh inserts any ListObject replacement, it releases
all selected sheets' old aliases inside that same transaction. A valid table
can therefore move from a later-order sheet to an earlier-order sheet without
colliding with its stale owner. Collisions with unchanged sheets or between
new owners still fail as `E_CORRUPT`, and any failure restores the old catalog.

R*Tree and interval backends retain equal source-refresh behavior. A schema-v2
derived sidecar rebuilds to v3 without reducing its monotonic generation.

## Pre-initial-freeze adversarial preflight

Non-verdict review and fuzzing were deliberately repeated before the initial
formal P3 freeze. The late passes found and permanently regressed:

- grouped whole-axis and name/structured endpoints whose outer colon had been
  lost;
- extra parentheses around named or lexical callable results;
- `@(...)` groups on both sides of one colon, which openpyxl had exposed as an
  empty synthetic function;
- exponential duplicate callable-choice expansion, now flattened,
  identity-deduplicated, capped at 32 alternatives, and conservatively marked
  reference-capable on overflow;
- mixed relative A1 plus fixed name/structured endpoints that were exact at the
  anchor but under-extruded over a block; endpoint components now extrude
  independently, with a deterministic 150-example two-dimensional/boundary
  differential against brute per-cell union;
- incremental table moves that collided with a stale alias; and
- R1C1 block-key collisions involving degenerate ranges and directional A1
  case spelling;
- formula/LAMBDA-name body precedents that were falsely promoted into exact
  colon hulls, including scalar functions and bare LAMBDAs;
- hidden `INDEX` results that lost `opaque:INDEX`/`I_DYNAMIC_REF` through LET
  or name aliases, and computed endpoints on whitespace intersections that
  lost all opacity; and
- parenthesized constant-name endpoints and valid computed-name alias chains,
  the latter previously growing exponentially through repeated expansion;
- whitespace-intersection scanning that stopped after the first operator,
  making later dynamic attribution depend on argument order;
- one nested computed-name colon endpoint that emitted two differently worded
  `I_DYNAMIC_REF` records for the same underlying `INDEX`; and
- an inherited mutable `ContextVar` cache shared by child tasks or copied
  thread contexts. The final owner key retains actual `Thread` and asyncio
  `Task` objects, so even a recycled Windows numeric thread id cannot reuse a
  prior execution's cache.

The 14-level duplicate-choice reproducer fell from about 3.68 seconds to about
0.0065 seconds. An execution-local, recursion-keyed name-expansion memo keeps
workbook contexts alive by identity and separates scopes, anchors, spellings,
spill modes, name stacks, tasks, and threads. A valid 32-name alias chain
completes in under 0.01 seconds on the development machine while the existing
depth guard still contains longer or recursive chains.
Independent post-fix matrices and the current corrected command results are
recorded below; these engineering preflights are not substituted for the
required formal verdicts or the later frozen fingerprint.

## Post-REVISE declaration and lexical-composite remediation

The initial mechanics reviewer challenged the declaration grammar by applying
the general Name Manager documentation to LET/LAMBDA locals. That claim was
not accepted on documentation alone. Microsoft Support's
[LET](https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999),
[LAMBDA](https://support.microsoft.com/en-us/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67),
and [formula-name](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64)
pages disagree at their boundaries, so the installed Excel runtime was treated
as the behavioral oracle.

On desktop Excel 16.0 build 19530 (64-bit), a 49-formula worksheet paste/entry
matrix produced 26 calculated formulas and 23 rejected formulas retained as
literal text. The exact UTF-8 rows are committed in
[`p3-excel-declaration-oracle.csv`](p3-excel-declaration-oracle.csv). The UI
accepted Unicode letters, underscore/backslash names, R1C1-like spellings, and
beyond-grid A1-like spellings; it rejected in-grid A1 spellings, periods,
operators, spaces, `@`/`#`, leading digits, and same-scope case-insensitive
duplicates. Nested case-insensitive shadowing remained valid. This is empirical
worksheet-entry evidence, not a claim about every Excel version.

COM `Range.Formula2` was probed separately and accepted some A1-looking local
names only in literal-valued expressions while rejecting their reference-valued
counterparts. It also confirmed that unprefixed built-ins win in function-call
position: same-spelled LET bindings did not shadow `SUM`, `INDEX`, `CHOOSE`,
`OFFSET`, or `IF`. Because this automation parser differs from worksheet entry,
it does not replace the UI matrix.

Fresh saved-workbook inspection exposed the actual implementation gap: Excel
serializes ordinary LET/LAMBDA declarations and their uses with `_xlpm.` (for
example `_xlfn.LET(_xlpm.x,$B$2,_xlpm.x)`). The corrected analyzer preserves
raw and stored namespaces as distinct exact keys, strips `_xlpm.` only for
display and duplicate comparison, suppresses prefixed locals, and never leaks
the private prefix through `function_calls` or dynamic labels. F19 and F20 now
use realistic stored spellings.

The final code preflight then found a separate ambiguity: an all-letter raw
local such as `r` in `r:s` was classified as a whole-column A1 endpoint before
lexical scope was consulted. The corrected analyzer gives an exact in-scope
LET/LAMBDA binding precedence over contextual whole-axis recovery, while known
defined names and ordinary unbound `R:S` references retain their existing
classification. Scalar locals remain non-reference-valued; reference-valued
locals retain their concrete precedents, dynamic attribution, and
`opaque:ref`. Qualified, structured, external, and 3-D spellings are excluded
from this recovery. A disposable Excel `Formula2` differential confirmed that
`LET(C,$B$2,SUM(C:C))` and its `R:R` equivalent use the lexical value, while
scalar-local variants produce Excel errors.

Independent read-only audits then found and regressed scope-blind intersection
preclassification, raw/stored or builtin/local collisions, direct and grouped
lexical colon/intersection endpoints, and repeated parenthesis-map construction.
The final 7,615-character/1,900-operand grouped-intersection probe fell from
1.18–1.60 seconds to 0.099 seconds by reusing the formula's matching-group map.
All fixes occurred after the initial frozen tree, so both formal domains must
inspect a new fingerprint.

## Fixture and oracle evidence

P3 adds F19 `modern_functions.xlsx` while preserving the locked bytes of the
P1 fixtures. F19 includes stored `_xlfn.`/`_xlws.` calls, `_xlpm.` LET/LAMBDA
locals, a function-position LAMBDA name, XLOOKUP, cell and name spills,
implicit intersection, a cached spill follower, and generator-computed formula
caches. P3 also corrects F20's generated LAMBDA definitions to use `_xlpm.` and
refreshes its exact workbook-map golden and 4,623-character/1,619-token budget.

The pinned-openpyxl cell-stream oracle has no active skips. F01, F02, F03,
F07, F12, F13, F14, F19, and F20 compare exactly, and repeated generation is
byte-identical. Parser regressions separately prove that modern shared groups
translate `@A2`, `A2:INDEX(...)`, `A2#`, `SUM(A2#)`, and an escaped structured
header without crashing or retaining the master's relative row.

Committed semantic goldens complement the cell-stream oracle:
`tests/golden/f07-formula-semantics.json` freezes the shared groups, tamper,
blocks, edges, and exact inconsistency diagnostic; and
`tests/golden/f19-formula-semantics.json` freezes the modern formulas, R1C1
blocks, spill/ref edges, flags, and empty diagnostic set. Their test compares
the natural pretty-JSON bytes and requires an explicit environment opt-in to
refresh either artifact.

P3 also intentionally changes the compact F03 workbook-map golden now that
formula semantics are populated: Calc reports four blocks, Summary reports
five, and the diagnostic summary contains one warning. `Calc!B2` is the
base-case formula and therefore the one-of-five minority in its maximal
vertical run under the frozen inconsistency threshold. The bounded map remains
exactly 1,041 serialized characters and 342 `o200k_base` tokens.

## Privacy and performance boundaries

Canonical exports and persisted edge labels are checked for URL credentials,
private path segments, query/fragment secrets, and raw external-link metadata.
A relationship-target-only mutation updates the private context and public
sanitized edge while reporting no reindexed worksheet.

The final exact 50,000-formula-column baseline uses the fill-down family
`=A<row>*$D$1` and three same-process repetitions. It measured medians of
2.218 seconds for R1C1 normalization, 0.097 seconds for guarded block
construction, and 0.226 seconds for inconsistency analysis on the resumed
2026-07-27 environment. A disposable
50,000-by-10 cold-index probe earlier in P3 measured 37.880 seconds; it is
engineering profile evidence, not the frozen P8 benchmark. P3 therefore does
**not** claim S1. Profiling attributes most remaining time to generic per-cell
OOXML parsing and three region projections, not block persistence. P8 owns the
frozen S1 gate; the recorded remediation path is row-end worksheet parsing with
a conservative common-cell fallback plus an exact dense-region accumulator,
with all current parser/oracle paths retained as differential references.

## Verification and review

Fresh corrected-candidate results on 2026-07-27:

| Check | Result |
|---|---|
| P3 formula/property slice | 501 passed in 16.64 seconds |
| Full suite with branch coverage | 771 passed in 95.41 seconds; 90.41% `excel_lsp.core` branch coverage (85% required) |
| Adversarial intersection stress | Four 33-intersection order/grouping variants each retained 20 `INDEX`, 8 named-LAMBDA, 1 generic marker, and 28 dynamic diagnostics |
| Inherited-context cache stress | 512 asyncio tasks plus 512 copied-context thread calls were isolated; an exact recycled-thread-id reproducer opened a fresh cache |
| Declaration runtime oracle | Exact 49-row worksheet-entry CSV validates as 26 accepted and 23 rejected; representative grammar, raw/stored namespace, and built-in precedence regressions are permanent |
| Lexical-composite stress | A 7,615-character formula with 1,900 grouped local operands completed in 0.099 seconds and retained all 1,899 intersection markers |
| Structured-context differential | A 101-case remediation preflight completed with 0 failures in 1.954 seconds; formal R-mech #12 independently expanded it to 605 block-versus-per-cell cases with 0 failures |
| Static gates | Ruff check clean; all 63 Python files format-clean; Pyright 0 errors/warnings/information; `git diff --check` clean |
| Reproducibility and package | Fixture generator completed; F19 semantic and F20 map goldens passed; `uv lock --check` resolved the frozen lock; sdist and wheel built successfully |
| Repository hygiene | 29 repository Markdown files and 37 local links checked with 0 missing; final tracked-file and forbidden-junk counts are fingerprint-time checks |

The corrected results above were produced on Windows 11, Python 3.12.11,
pytest 9.1.1, and Excel 16.0 build 19530 where a desktop oracle was required.
The exact verification commands were:

```text
uv run pytest tests/unit/test_reference_extraction.py tests/unit/test_formula_translation.py tests/unit/test_formula_blocks.py tests/unit/test_formula_indexing.py tests/unit/test_formula_index_integration.py tests/unit/test_p3_semantic_goldens.py tests/unit/test_p3_storage_foundation.py tests/property/test_r1c1_properties.py tests/property/test_mixed_endpoint_extrusion_properties.py tests/property/test_structured_context_properties.py tests/unit/test_fixture_generation.py tests/oracle/test_oracle.py -q
uv run pytest --cov=excel_lsp.core --cov-branch --cov-report=term-missing
uv run ruff check .
uv run ruff format --check .
uv run pyright
uv run python tests/fixtures/generate.py
uv lock --check
uv build
git diff --check
```

### Initial formal split and reopened gate

The first frozen candidate was bound to:

- base `3359a974dc72db1dd1ec47507eaf24891c670c92`;
- stage tree `4b8364b6060ec52245260448601c9c61f19856a3`;
- cached-diff blob `1a0b6eb4c92cc56345203f1a7f4171aab7c66628`;
- 45 files, 14,600 insertions, and 156 deletions.

Global review invocation #17 (R-mech #10) returned `REVISE` over declaration
semantics. Global invocation #18 (R-test #5) returned a clean `APPROVE` after
405 focused tests, all 644 repository tests, and 90.37% core branch coverage.
That test verdict proves only the exact tree above. The declaration/runtime
investigation reopened and changed the candidate, so neither old fingerprint
nor old test verdict closes P3.

That tree did not close P3. Its corrected successor and fresh verdicts are
recorded below.

### Second formal split and contextual-tiling remediation

The second frozen candidate was bound to:

- base `3359a974dc72db1dd1ec47507eaf24891c670c92`;
- stage tree `c7de5e88dbccc9b0c0f7276af1ed3bd585d8ee0a`;
- cached-diff blob `f4e58e1213c4b4616a3a8d9e05d6c05255df9eeb`;
- 51 files, 15,923 insertions, and 193 deletions.

Global invocation #19 (R-mech #11) returned `REVISE` with two majors. A
compound reference such as `[@Input]:F2` had lost its current-table requirement
after endpoint references were retokened to the whole composite, and semantic
tiles analyzed the block-anchor formula at a new coordinate without first
translating relative A1 endpoints. Global invocation #20 (R-test #6) returned a
clean `APPROVE` on that same tree after independently reproducing 494 focused
tests, all 764 repository tests, and 90.35% branch coverage. That approval is
exact-tree evidence only and was invalidated by the mechanics correction.

The corrected implementation recursively aggregates structured-context
requirements across safe range/intersection/union operands, including multiple
qualified current-row tables. Each semantic tile translates the block-anchor
formula to the tile anchor with the shared modern-safe translator while
preserving coordinate spill anchors. Follow-up block-versus-per-cell auditing
then exposed two context-boundary provenance gaps: grouped colon endpoints
could change between static and contextual ownership, and an intersection that
became valid only in a later tile could lose `opaque:ref`. Opaque structured
endpoints now remain foldable composite atoms, and tile analysis unions the
context-induced opacity marker. Permanent tests cover horizontal and vertical
tiling, both directions across table boundaries, direct/reversed/grouped
spelling, multiple qualified tables, intersections, initial-anchor recovery,
and fixed spill anchors.

### Final frozen candidate and closed gate

The post-remediation candidate was bound to:

- base `3359a974dc72db1dd1ec47507eaf24891c670c92`;
- stage tree `a6ced2a07ca3e438c47e11e93e261b572699cc50`;
- cached-diff blob `fe3d0c40a75446da3d1b6af146292eae778eb3f6`;
- 51 files, 16,309 insertions, and 194 deletions;
- 107 tracked files and zero tracked forbidden-junk paths.

Global invocation #21 (R-mech #12) returned a clean `APPROVE` with no findings
after independently reproducing 501 focused tests, all 771 repository tests,
90.41% branch coverage, and a 605-case block-versus-per-cell structured-context
differential with zero failures. Global invocation #22 (R-test #7) returned a
clean `APPROVE` with no findings on the same tree after independently
reproducing the 501/771 test totals and 90.41% coverage. Both reviewers verified
the stage tree and cached-diff blob before and after review, with no unstaged or
untracked changes. P3's mechanics and test gates are therefore closed on this
exact fingerprint.
