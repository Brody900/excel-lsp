# README claims-to-artifacts plan

This is the frozen bridge between the P2 README skeleton and the evidence that
later phases must produce before v0.1.0. It prevents persuasive copy from
outrunning implementation or measured results.

Status meanings:

- **Verified Pn**: the cited proof exists from a completed phase.
- **Planned Pn**: the README must retain an explicit future-status label until
  the named phase commits and verifies every listed artifact.
- **Scope declaration**: a product boundary or roadmap statement, not an
  empirical performance claim.

Every entry in the artifact column is an exact repository path, Markdown
anchor, or pytest node ID. Descriptive placeholders such as “matching SVG,”
“the relevant tests,” or “raw results” are intentionally forbidden. P9 must
audit every row and must not promote `Planned` to `Verified` without
following all listed references.

| ID | README item or claim | Phase | Status | Required committed artifact |
|---|---|---:|---|---|
| P1-FOUND | The streaming parser, persistent SQLite index, and incremental freshness foundation are implemented. | P1 | Verified P1 | `docs/evidence/p1-foundation.md#delivered-contracts`<br>`tests/unit/test_ooxml_parser.py::test_streams_all_cell_types_formulas_and_actual_bounds`<br>`tests/unit/test_index_lifecycle.py::test_incremental_and_fresh_full_indexes_have_equal_canonical_exports`<br>`tests/unit/test_index_store.py::test_store_configures_sqlite_and_creates_frozen_schema` |
| P2-FOUND | Sparse regions, frozen public symbol IDs, and the compact workbook map are implemented and tested. | P2 | Verified P2 | `docs/evidence/p2-regions-map.md#delivered-contracts`<br>`tests/unit/test_regions.py::test_listobject_is_exact_authoritative_and_excludes_totals_from_profiles`<br>`tests/unit/test_symbols.py::test_frozen_symbol_ids_match_the_handoff_scheme`<br>`tests/unit/test_workbook_map.py::test_workbook_map_uses_stable_ids_caps_content_and_surfaces_visibility` |
| POS-01 | Python core, CLI, and stdio MCP server support modern OOXML workbooks. | P7/P9 | Planned P9 | `docs/architecture.md#layer-boundaries`<br>`docs/evidence/p7-mcp-cli.md#package-and-transport-boundary`<br>`docs/evidence/fresh-install.md#wheel-install` |
| POS-02 | LSP-style means symbols, references, diagnostics, and incremental indexing rather than the LSP wire protocol. | P2 | Verified P2 | `README.md`<br>`tests/unit/test_readme_contract.py::test_readme_has_frozen_positioning_and_section_order` |
| POS-03 | Agents navigate semantically rather than dumping complete workbooks into context. | P7/P8 | Planned P8 | `tests/mcp/test_conformance.py::test_navigation_flow_never_returns_bulk_workbook_data`<br>`benchmarks/results/scripted.csv`<br>`benchmarks/results/llm-eval.jsonl` |
| S1 | A 50,000-row by 10-column workbook indexes in under 10 seconds cold and under 1 second after a one-sheet change. | P8 | Planned P8 | `benchmarks/results/index-timing.csv`<br>`docs/assets/benchmark-index-time.png`<br>`docs/assets/benchmark-index-time.svg`<br>`docs/evidence/success-criteria.md#s1` |
| S2 | The F03 `open_workbook` map is at most 1,500 `o200k_base` tokens. | P2 | Verified P2 | `tests/golden/f03-workbook-map.json`<br>`benchmarks/results/map-budgets.json`<br>`docs/evidence/p2-regions-map.md#map-budgets`<br>`tests/unit/test_region_index_integration.py::test_workbook_map_matches_golden_and_budget` |
| S2-CAP | Every map, including F20, fits 8,000 serialized characters with deterministic degradation. | P2 | Verified P2 | `tests/golden/f20-workbook-map.json`<br>`benchmarks/results/map-budgets.json`<br>`docs/evidence/p2-regions-map.md#map-budgets`<br>`tests/unit/test_region_index_integration.py::test_workbook_map_matches_golden_and_budget`<br>`tests/unit/test_region_index_integration.py::test_every_current_fixture_map_obeys_the_character_cap` |
| S3 | Cross-sheet traces are exact on F03, F04, F05, F15, and F19. | P4 | Planned P4 | `tests/golden/f03-traces.json`<br>`tests/golden/f04-traces.json`<br>`tests/golden/f05-traces.json`<br>`tests/golden/f15-traces.json`<br>`tests/golden/f19-traces.json`<br>`docs/evidence/p4-graph.md#cross-sheet-trace-matrix` |
| S4 | Edited workbooks open without repair, recalculate correctly, and preserve every untouched ZIP part byte-identically. | P6/P8 | Planned P8 | `docs/evidence/part-diff-f16.json`<br>`docs/evidence/part-diff-f21.json`<br>`docs/evidence/live-excel/index.md#round-trip-results`<br>`docs/evidence/live-excel/03-no-repair.png`<br>`docs/evidence/live-excel/04-vba-stamp.png`<br>`docs/evidence/live-excel/06-chart-intact.png` |
| S5 | Benchmarks show ≥ 10× token reduction vs. the naive-dump baseline on the defined task suite, with equal-or-better task accuracy in LLM evals. | P8 | Planned P8 | `benchmarks/results/scripted.csv`<br>`benchmarks/results/llm-eval.jsonl`<br>`benchmarks/results/accuracy.csv`<br>`benchmarks/check.py`<br>`docs/evidence/success-criteria.md#s5` |
| S6 | No tool response exceeds 200 values or 8,000 serialized characters. | P7 | Planned P7 | `tests/mcp/test_conformance.py::test_all_tool_responses_obey_value_and_character_caps`<br>`docs/evidence/p7-mcp-cli.md#response-caps` |
| S7 | `uvx excel-lsp serve`, or the documented git fallback, works from a clean environment. | P9 | Planned P9 | `docs/evidence/fresh-install.md#uvx-install`<br>`docs/evidence/fresh-install.md#git-fallback`<br>`docs/evidence/success-criteria.md#s7` |
| HERO-01 | The top-of-README grouped, log-scale token chart reflects checked benchmark rows. | P8 | Planned P8 | `docs/assets/benchmark-token-hero.png`<br>`docs/assets/benchmark-token-hero.svg`<br>`benchmarks/results/scripted.csv`<br>`benchmarks/results/llm-eval.jsonl`<br>`benchmarks/plot.py` |
| DEMO-01 | A 60-second lineage demo shows the verified workflow. | P8 | Planned P8 | `docs/assets/lineage-demo.gif`<br>`docs/evidence/live-excel/index.md#lineage-demo`<br>`docs/evidence/live-excel/demo-capture.json` |
| QS-01 | `uvx excel-lsp serve` starts the release server. | P7/P9 | Planned P9 | `docs/evidence/fresh-install.md#uvx-install`<br>`tests/mcp/test_conformance.py::test_initialize_and_list_tools_over_installed_stdio_server` |
| QS-02 | `codex mcp add excel-lsp -- uvx excel-lsp serve` is current Codex syntax. | P9 | Planned P9 | `docs/evidence/codex-mcp-help.txt`<br>`docs/evidence/fresh-install.md#codex-mcp-registration` |
| QS-03 | The Codex TOML example launches the same stdio server. | P7/P9 | Planned P9 | `examples/codex.config.toml`<br>`docs/evidence/fresh-install.md#codex-toml-configuration` |
| QS-04 | The generic `.mcp.json` example is syntactically valid for clients that use that format. | P7 | Planned P7 | `examples/mcp.json`<br>`tests/unit/test_readme_contract.py::test_codex_and_generic_mcp_examples_are_equivalent` |
| TOOL-01 | `open_workbook` contract. | P7 | Planned P7 | `docs/tool-reference.md#open_workbook`<br>`tests/mcp/test_conformance.py::test_open_workbook_happy_and_error_contract` |
| TOOL-02 | `refresh` contract. | P7 | Planned P7 | `docs/tool-reference.md#refresh`<br>`tests/mcp/test_conformance.py::test_refresh_happy_and_error_contract` |
| TOOL-03 | `list_symbols` contract. | P7 | Planned P7 | `docs/tool-reference.md#list_symbols`<br>`tests/mcp/test_conformance.py::test_list_symbols_happy_and_error_contract` |
| TOOL-04 | `get_region_schema` contract and bounded samples. | P7 | Planned P7 | `docs/tool-reference.md#get_region_schema`<br>`tests/mcp/test_conformance.py::test_get_region_schema_happy_error_and_wide_sample_cap` |
| TOOL-05 | `read_range` pagination and 200-value limit. | P7 | Planned P7 | `docs/tool-reference.md#read_range`<br>`tests/mcp/test_conformance.py::test_read_range_pagination_value_cap_and_stale_cursor` |
| TOOL-06 | `find` bounded regex search. | P7 | Planned P7 | `docs/tool-reference.md#find`<br>`tests/mcp/test_conformance.py::test_find_happy_error_timeout_and_snippet_cap` |
| TOOL-07 | `trace_precedents` contract. | P7 | Planned P7 | `docs/tool-reference.md#trace_precedents`<br>`tests/mcp/test_conformance.py::test_trace_precedents_happy_and_error_contract` |
| TOOL-08 | `trace_dependents` contract. | P7 | Planned P7 | `docs/tool-reference.md#trace_dependents`<br>`tests/mcp/test_conformance.py::test_trace_dependents_happy_and_error_contract` |
| TOOL-09 | `trace_path` contract. | P7 | Planned P7 | `docs/tool-reference.md#trace_path`<br>`tests/mcp/test_conformance.py::test_trace_path_connected_unconnected_depth_and_path_caps` |
| TOOL-10 | `explain_formula` contract. | P7 | Planned P7 | `docs/tool-reference.md#explain_formula`<br>`tests/mcp/test_conformance.py::test_explain_formula_happy_and_error_contract` |
| TOOL-11 | `get_diagnostics` filtering contract. | P7 | Planned P7 | `docs/tool-reference.md#get_diagnostics`<br>`tests/mcp/test_conformance.py::test_get_diagnostics_filters_counts_and_error_contract` |
| TOOL-12 | `profile_column` bounded statistics contract. | P7 | Planned P7 | `docs/tool-reference.md#profile_column`<br>`tests/mcp/test_conformance.py::test_profile_column_numeric_text_and_missing_cache_contract` |
| TOOL-13 | `write_cells` contract and destructive annotation. | P7 | Planned P7 | `docs/tool-reference.md#write_cells`<br>`tests/mcp/test_conformance.py::test_write_cells_happy_error_and_destructive_annotation` |
| TOOL-14 | `set_column_formula` contract and destructive annotation. | P7 | Planned P7 | `docs/tool-reference.md#set_column_formula`<br>`tests/mcp/test_conformance.py::test_set_column_formula_happy_error_and_destructive_annotation` |
| TOOLS-15 | Exactly 14 tools comprise 12 reads and 2 destructive writes. | P7 | Planned P7 | `tests/mcp/test_conformance.py::test_lists_exactly_fourteen_tools_with_annotations_and_instructions`<br>`docs/evidence/p7-mcp-cli.md#tool-inventory` |
| ARCH-01 | Loader to SQLite and spatial index to graph to MCP architecture. | P1/P4/P7 | Planned P7 | `docs/architecture.md#verified-p1-data-flow`<br>`docs/evidence/p1-foundation.md#delivered-contracts`<br>`docs/evidence/p4-graph.md#architecture`<br>`docs/evidence/p7-mcp-cli.md#package-and-transport-boundary` |
| ARCH-02 | Core remains usable without the MCP server. | P7 | Planned P7 | `tests/unit/test_package.py::test_core_layer_is_importable`<br>`tests/unit/test_cli.py::test_cli_imports_core_without_server_startup`<br>`docs/evidence/p7-mcp-cli.md#core-embedding-boundary` |
| BENCH-01 | Six deterministic tasks use exact final `ANSWER:` JSON grading. | P8 | Planned P8 | `benchmarks/tasks/B1.md`<br>`benchmarks/tasks/B2.md`<br>`benchmarks/tasks/B3.md`<br>`benchmarks/tasks/B4.md`<br>`benchmarks/tasks/B5.md`<br>`benchmarks/tasks/B6.md`<br>`benchmarks/check.py`<br>`tests/unit/test_benchmark_checkers.py` |
| BENCH-02 | Scripted and headless-Codex measurement methods are reproducible. | P8 | Planned P8 | `benchmarks/run_scripted.py`<br>`benchmarks/run_llm_eval.py`<br>`benchmarks/results/environment.json`<br>`benchmarks/README.md#measurement-modes`<br>`docs/evidence/codex-exec-help.txt` |
| BENCH-CLI | `excel-lsp bench` runs the reproducible benchmark harness advertised in the README. | P7/P8 | Planned P8 | `tests/unit/test_cli.py::test_bench_command_runs_reproducible_harness`<br>`docs/evidence/p8-benchmarks.md#excel-lsp-bench`<br>`benchmarks/run_scripted.py`<br>`benchmarks/check.py` |
| BENCH-03 | Both headless-Codex repetitions and their agreement are visible. | P8 | Planned P8 | `benchmarks/results/llm-eval.jsonl`<br>`benchmarks/results/accuracy.csv`<br>`benchmarks/README.md#llm-repetitions-and-agreement` |
| BENCH-04 | Five rendered benchmark assets have matching source data and scripts. | P8 | Planned P8 | `docs/assets/benchmark-token-hero.png`<br>`docs/assets/benchmark-token-hero.svg`<br>`docs/assets/benchmark-token-modes.png`<br>`docs/assets/benchmark-token-modes.svg`<br>`docs/assets/benchmark-tool-calls.png`<br>`docs/assets/benchmark-tool-calls.svg`<br>`docs/assets/benchmark-index-time.png`<br>`docs/assets/benchmark-index-time.svg`<br>`docs/assets/benchmark-audit-cost.png`<br>`docs/assets/benchmark-audit-cost.svg`<br>`benchmarks/results/scripted.csv`<br>`benchmarks/results/llm-eval.jsonl`<br>`benchmarks/results/index-timing.csv`<br>`benchmarks/results/audit-cost.json`<br>`benchmarks/plot.py` |
| COMP-01 | Persistent semantic index comparison row. | P9 | Planned P9 | `docs/evidence/comparison-sources.md#persistent-semantic-index`<br>`docs/evidence/p1-foundation.md#delivered-contracts`<br>`benchmarks/baseline_server.py` |
| COMP-02 | Formula dependency graph comparison row. | P9 | Planned P9 | `docs/evidence/comparison-sources.md#formula-dependency-graph`<br>`docs/evidence/p4-graph.md#graph-query-evidence` |
| COMP-03 | Incremental reindex comparison row. | P9 | Planned P9 | `docs/evidence/comparison-sources.md#incremental-reindex`<br>`docs/evidence/p1-foundation.md#invariant-evidence`<br>`benchmarks/results/index-timing.csv` |
| COMP-04 | Formula diagnostics comparison row. | P9 | Planned P9 | `docs/evidence/comparison-sources.md#formula-diagnostics`<br>`docs/evidence/p5-diagnostics.md#diagnostic-matrix` |
| COMP-05 | Edit support and untouched-part fidelity comparison row. | P9 | Planned P9 | `docs/evidence/comparison-sources.md#edit-support-and-fidelity`<br>`docs/evidence/part-diff-f16.json`<br>`docs/evidence/part-diff-f21.json`<br>`docs/evidence/live-excel/index.md#round-trip-results` |
| COMP-06 | Token-discipline comparison row. | P8/P9 | Planned P9 | `docs/evidence/comparison-sources.md#token-discipline`<br>`benchmarks/baseline_server.py`<br>`tests/mcp/test_conformance.py::test_all_tool_responses_obey_value_and_character_caps`<br>`benchmarks/results/scripted.csv` |
| HOW-01 | Sparse, ListObject-first region detection exposes confidence. | P2 | Verified P2 | `docs/evidence/p2-regions-map.md#delivered-contracts`<br>`docs/index-internals.md#p2-regions-and-headers-contract`<br>`tests/unit/test_regions.py::test_listobject_is_exact_authoritative_and_excludes_totals_from_profiles`<br>`tests/unit/test_regions.py::test_two_row_merged_headers_are_synthesized_and_win_the_score`<br>`tests/property/test_region_properties.py::test_i4_random_sparse_cells_belong_to_exactly_one_nonoverlapping_region`<br>`tests/property/test_region_properties.py::test_i5_listobject_range_is_exact_and_heuristics_never_intersect_it`<br>`tests/property/test_region_properties.py::test_i6_identical_content_has_identical_ordinals_across_stream_chunking` |
| HOW-02 | R1C1 normalization forms formula blocks. | P3 | Planned P3 | `docs/evidence/p3-formulas-blocks.md#r1c1-and-block-evidence`<br>`docs/index-internals.md#formula-blocks`<br>`tests/property/test_r1c1_properties.py::test_i10_translation_invariance` |
| HOW-03 | RTree range-edge storage has an interval fallback. | P4 | Planned P4 | `docs/evidence/p4-graph.md#spatial-edge-evidence`<br>`docs/index-internals.md#range-edges`<br>`tests/unit/test_graph.py::test_rtree_and_interval_backends_have_equal_graph_queries` |
| SEC-01 | Runtime operates on local files and makes no network requests. | P7/P9 | Planned P9 | `SECURITY.md#runtime-data-handling`<br>`tests/mcp/test_conformance.py::test_server_operates_with_network_denied`<br>`docs/evidence/fresh-install.md#runtime-network-audit` |
| SEC-02 | The two write tools modify only deliberate worksheet and calculation parts. | P6 | Planned P6 | `docs/evidence/part-diff-f16.json`<br>`docs/evidence/part-diff-f21.json`<br>`tests/property/test_editor_preservation.py::test_i18_random_edit_scripts_preserve_untouched_parts`<br>`docs/evidence/p6-editor.md#part-preservation` |
| SEC-03 | `EXCEL_LSP_ROOT` applies realpath confinement and defaults to unrestricted. | P7 | Planned P7 | `SECURITY.md#path-access`<br>`tests/mcp/test_conformance.py::test_excel_lsp_root_allows_denies_and_resolves_symlinks`<br>`docs/evidence/p7-mcp-cli.md#path-confinement` |
| LIM-01 | There is no formula engine; cached values are read and Excel recalculates. | P6/P8 | Planned P8 | `tests/unit/test_profile_column.py::test_missing_caches_return_recalculation_hint`<br>`docs/evidence/p6-editor.md#recalculation-boundary`<br>`docs/evidence/live-excel/index.md#recalculation-results` |
| LIM-02 | Dynamic references such as `INDIRECT` are opaque but diagnosed. | P3/P5 | Planned P5 | `tests/unit/test_reference_extraction.py::test_dynamic_functions_retain_explicit_refs_and_emit_opaque_edges`<br>`tests/golden/f11-diagnostics.json`<br>`docs/evidence/p5-diagnostics.md#dynamic-reference-diagnostics` |
| LIM-03 | Header inference is heuristic and exposes confidence. | P2 | Verified P2 | `docs/evidence/p2-regions-map.md#header-scoring-decision`<br>`tests/unit/test_regions.py::test_numeric_grid_has_no_inferred_header_and_invalid_styles_are_safe`<br>`tests/unit/test_regions.py::test_two_row_merged_headers_are_synthesized_and_win_the_score` |
| LIM-04 | Writes use inline strings, which some third-party tools handle poorly. | P6 | Planned P6 | `tests/unit/test_editor.py::test_string_writes_use_inline_strings`<br>`docs/evidence/p6-editor.md#inline-string-compatibility` |
| LIM-05 | Datetime writes are rejected in v0.1.0. | P6/P7 | Planned P7 | `tests/unit/test_editor.py::test_datetime_write_returns_invalid_value`<br>`tests/mcp/test_conformance.py::test_write_cells_rejects_datetime_without_traceback` |
| LIM-06 | Writes inside multi-cell array formulas are refused. | P6/P7 | Planned P7 | `tests/unit/test_editor.py::test_multicell_array_edit_returns_array_formula_error`<br>`tests/mcp/test_conformance.py::test_write_tools_refuse_multicell_array_members` |
| LIM-07 | Spill extents are not statically tracked. | P3 | Planned P3 | `tests/unit/test_formula_indexing.py::test_spill_edge_resolves_only_the_anchor_instead_of_extruding_the_block`<br>`tests/unit/test_formula_index_integration.py::test_f19_persists_modern_edges_without_spurious_diagnostics`<br>`docs/evidence/p3-formulas-blocks.md#spill-reference-scope` |
| NGOAL-01 | No chart or pivot creation, rename, Sheets, binary Excel, collaborative editing, runtime network, or telemetry in v0.1.0. | P2/P9 | Scope declaration | `README.md#non-goals-for-v010`<br>`SECURITY.md#runtime-data-handling`<br>`tests/unit/test_phase1_edge_cases.py::test_index_workbook_uses_structured_path_errors` |
| ROAD-01 | Rename refactoring is the flagship v1.x roadmap item. | P9 | Scope declaration | `README.md#roadmap`<br>`CHANGELOG.md#planned-before-v010` |
| ROAD-02 | Real LSP, multi-workbook graph, diff, watch, binary Excel, Sheets, datetime writes, and stable aliases remain roadmap work. | P9 | Scope declaration | `README.md#roadmap`<br>`CHANGELOG.md#planned-before-v010` |
| EVID-01 | The evidence section links every completed phase and S1 through S7. | P9 | Planned P9 | `docs/evidence/README.md`<br>`docs/evidence/success-criteria.md`<br>`tests/unit/test_documentation_links.py::test_all_repository_markdown_links_and_anchors_resolve` |
| LEGAL-01 | Microsoft non-affiliation and trademark footer uses the frozen wording. | P2 | Verified P2 | `README.md`<br>`tests/unit/test_readme_contract.py::test_readme_has_frozen_positioning_and_section_order` |

## P2 repository-review rule

The first required early P2 R-repo invocation is charged as `REVISE`. Under the
user-authorized pooled-review amendment, a fresh stateless R-repo re-review
follows each `REVISE` until the P2 repository gate approves or the overall
review pool reaches its documented exhaustion policy. Before every invocation,
the orchestrator must run the README contract test, local-link and anchor
audit, junk-file audit, Markdown review, map-budget integration test, and
ordinary repository checks. The reviewer packet must quote the P2-specific
scope from the handoff: this early gate approves the README skeleton and
exhaustive claims-to-artifacts plan, not the deliberately future P7-P9 release
evidence.

The reviewer should reject misleading present-tense copy, a missing P2 proof,
or an unmapped claim. Later-phase rows plainly marked `Planned` are not
assertions that those features or measurements exist today. The eventual P9
audit must replace every release-relevant `Planned` status with a
verified status or remove the unsupported public claim. It must not silently
weaken a success criterion to fit observed results.

Current accounting: the first two P2 R-repo invocations remain charged as
`REVISE`; fresh invocation #3 returned `APPROVE` with one current-status
documentation minor corrected in this candidate. The early P2 repository gate
is approved. The separate minimum reserve of three R-repo invocations for P9
remains unchanged.
