# P5 diagnostics evidence

P5 completes the shared diagnostics catalog, persists the findings whose input
data exists by this phase, and exposes deterministic core filtering for the P7
MCP tool. One exact frozen fingerprint received both formal approvals.

## Diagnostic matrix

| Code | Severity | Producer and current evidence |
|---|---|---|
| `E_ERRVAL` | error | P5 scans stored cells by `value_type = 'error'`; F08 proves all specified values plus unrecognized `#FIELD!`. |
| `E_CIRCULAR` | error | P4 exact circular verifier; retained unchanged and returned through the P5 query model. |
| `E_BROKEN_XLINK` | error or warn | P5 resolves local targets relative to the workbook directory; missing local files are errors, while remote or unsafe-to-check targets are warnings. |
| `W_POSSIBLE_CIRCULAR` | warn | P4 bounded circular fallback; retained unchanged and returned through the P5 query model. |
| `W_INCONSISTENT_FORMULA` | warn | P3 minority-pattern detector, including F07's tamper. |
| `W_UNKNOWN_NAME` | warn | P3 formula analysis with LET/LAMBDA suppression. |
| `W_PARSE` | warn | P3 bounded formula-analysis fallback. |
| `W_LARGE_SHEET` | warn | P2 sparse-region analysis; formula refresh preserves it. |
| `W_REGEX_TIMEOUT` | warn | Typed P5 constructor is complete; the regex guard invokes it when P7 implements `find`. |
| `I_DYNAMIC_REF` | info | P3 dynamic-reference analysis; F11 freezes `INDIRECT` and `OFFSET`. |
| `I_VOLATILE` | info | P5 emits one finding per volatile formula block; F11 and F18 cover dynamic and non-dynamic volatility. |
| `I_STALE` | info | Typed P5 rectangle constructor is complete; P6 owns staleness production and clearing. |

The later-phase ownership in the deferred producer rows is deliberate: P5
defines and tests the shared diagnostic shape, while P6/P7 supply the state and
runtime deadline that do not exist yet. No placeholder row is persisted.

## Cached-error evidence

F08's label column contains ordinary strings such as `Expected #REF!`; its ten
formula-result cells carry OOXML `t="e"`. The result is exactly ten
`E_ERRVAL` findings, all on column B. The committed golden includes the nine
message-prettification values from HANDOFF section 5.6 and `#FIELD!`, which is
not in that list. The latter still produces `E_ERRVAL`, proving that the cell
type—not a text whitelist—controls detection.

Artifacts:

- `tests/fixtures/generate.py::_generate_f08`
- `tests/unit/test_diagnostics.py::test_f08_uses_ooxml_error_type_not_a_text_whitelist`
- `tests/golden/p5-diagnostics.json`

## External-link evidence

F10 contains a genuine workbook `<externalReferences>` entry, internal link
part, external relationship, content-type override, and `[1]Data!A1` formula.
The missing relative target resolves from the workbook directory and yields one
error on the source sheet. Separate tests prove an existing sibling file yields
no finding, HTTP(S) and unsafe paths warn rather than claim absence,
and credentials, directories, query strings, fragments, and tokens never enter
the diagnostic message or `related` data. UNC paths and authority-bearing
`file:` URIs are classified before any filesystem call, so link health never
contacts a network share. A target-only workbook refresh replaces both the
external edge label and the diagnostic.

Link health is also part of lifecycle freshness independently of workbook
bytes. Creating or deleting F10's target while leaving the workbook untouched
recomputes only workbook-wide external diagnostics, increments generation, and
returns no reindexed sheets. Two numeric M1 links with the same redacted
`budget.xlsx` basename remain associated with their exact source sheets by
reanalyzing only persisted external source blocks and retaining the `[n]`
identity internally; public data stays redacted.

Artifacts:

- `tests/fixtures/generate.py::inject_external_link`
- `tests/unit/test_diagnostics.py::test_external_link_health_distinguishes_existing_missing_remote_and_unsafe`
- `tests/unit/test_diagnostics.py::test_network_external_targets_warn_without_filesystem_io`
- `tests/unit/test_diagnostics.py::test_unchanged_workbook_refresh_tracks_external_target_create_and_delete`
- `tests/unit/test_diagnostics.py::test_numeric_external_link_identity_survives_same_basename_redaction`
- `tests/unit/test_formula_index_integration.py::test_external_target_only_refresh_updates_raw_context_and_persisted_edge`

## Dynamic-reference diagnostics

F11 persists `I_DYNAMIC_REF` and `I_VOLATILE` for independent `INDIRECT` and
`OFFSET` blocks. F18 independently freezes `NOW` and `RAND` volatility. A
50-cell contiguous `NOW()` fill produces exactly one block-level finding,
preventing a per-cell diagnostic explosion.

Artifacts:

- `tests/unit/test_diagnostics.py::test_f11_and_f18_emit_one_volatile_finding_per_block`
- `tests/unit/test_diagnostics.py::test_contiguous_volatile_fill_is_diagnosed_once_at_block_granularity`
- `tests/golden/p5-diagnostics.json`

## Filtering and lifecycle evidence

`IndexStore.get_diagnostics` filters before its 100-row cap by exact sheet,
severity, and catalog code. Counts are computed before limiting; deterministic
ordering is error, warning, info, then workbook sheet order and location.
Returned diagnostics and count mappings are immutable. Unknown filters are
rejected, while malformed stored code, severity, coordinates, or JSON is
shaped as `E_CORRUPT` rather than leaking an implementation exception. The
entire filtered cursor is validated before only the first 100 records are
materialized, so corruption at row 101 cannot hide behind the public cap.
Related values reject non-finite floats and container nesting beyond 64 levels,
preserving strict JSON shape without risking recursion failure. The persisted
boundary also converts decoder/freezer `RecursionError` to `E_CORRUPT`.

Incremental replacement deletes and recreates only P5 findings owned by the
selected sheet; unrelated sheet findings survive. The external-link set is
recomputed workbook-wide, because link metadata is workbook context. The
`diagnostic_analysis_version` participates in both the fast path and full-build
decision so an older sidecar cannot silently retain an incomplete catalog.

Artifacts:

- `tests/unit/test_diagnostics.py::test_get_diagnostics_filters_counts_caps_and_immutability`
- `tests/unit/test_diagnostics.py::test_selected_sheet_refresh_replaces_only_its_p5_diagnostics`
- `tests/unit/test_diagnostics.py::test_get_diagnostics_shapes_corrupt_rows`
- `tests/unit/test_diagnostics.py::test_get_diagnostics_validates_corruption_beyond_the_materialized_page`
- `tests/unit/test_diagnostics.py::test_get_diagnostics_shapes_excessive_nesting_beyond_the_materialized_page`
- `tests/unit/test_index_lifecycle.py::test_missing_diagnostic_analysis_version_forces_full_semantic_refresh`

## Fixture and oracle evidence

P5 adds deterministic F08, F10, F11, and F18 workbooks after generating the
historical corpus, preserving every earlier frozen hash. Their own SHA-256
values are locked, two complete generations are byte-identical, and all four
match the independent openpyxl dual-load oracle with no skip-list entry.

The current corpus contains 18 of the handoff's 21 generated fixture IDs. F06
belongs to P8 performance work; F16 and F21 belong to P6 editor-preservation
work. The fixture README and oracle skip list state that boundary explicitly.

## Phase verification

Fresh results on the implementation candidate:

| Check | Result |
|---|---|
| Full repository tests | 1,973 passed in 203.94 seconds |
| Full branch-coverage run | 1,973 passed in 374.51 seconds; 90.01% total core coverage and 92% for `core/diagnostics.py` |
| P5 diagnostic unit module | 29 passed; 92% branch coverage in the full instrumented run |
| Remediated diagnostics/lifecycle/formula/fixture/oracle/circular set | 127 passed in 31.53 seconds |
| Fixture shape + oracle set | 26 passed in 15.63 seconds |
| Pyright | 0 errors, 0 warnings |
| Ruff lint and format | passed; 76 files already formatted |
| Deterministic fixture regeneration | passed; all 18 current generated fixture IDs emitted with no tracked drift |
| Lock and distribution build | `uv lock --check` resolved 70 packages; sdist and wheel built successfully |
| Git whitespace check | passed |

The extra author-loop checks bound arbitrary error text so one malformed cached
error cannot dominate a response, deeply freeze structured `related` data, and
reject internally inconsistent aggregate counts. Formal #51/#52 then
reproduced four boundary failures. Main-agent remediation added the exact
unchanged-workbook health toggle, network no-I/O, numeric same-basename
ownership, post-cap corruption, and non-finite JSON regressions. Those
dimensions protect reproduced behavior. Formal #53/#54 then reproduced an
excessive-nesting recursion escape; the 2,000-level payload now has its own
beyond-page `E_CORRUPT` regression and a 64-container constructor bound. No
combinatorial matrix was added.

The approved frozen fingerprint was based on
`2024813fc6b34264212a5f89903b5f8391b2030b`, with staged tree
`bf7e9f2787df2110ae7b59f7aee7b82300dddc88` and cached binary-diff hash
`da558fb07f4320cbca1bd8876b35792baca46ce8`.

## Formal phase gate

The first combined reviewer returned R-mech #27 and R-test #22 as `REVISE` on
tree `ce38f057ff807a5c753807ad990dfb3784ded1c8`. Remediation and broad checks
produced tree `7a2fae24cb069931da2c82ccdef3f619b63380d4`; the same reviewer returned
R-mech #28 and R-test #23 as `REVISE` only for the nested-related recursion
escape. After remediation and broad checks, the same reviewer evaluated tree
`bf7e9f2787df2110ae7b59f7aee7b82300dddc88` and returned clean
`R-mech: APPROVE` (#29) and `R-test: APPROVE` (#24) verdicts. It verified the
base, tree, and cached-diff fingerprint at entry and exit, reran the 41-test
focused slice plus Ruff, formatting, and Pyright, and reported no findings.
