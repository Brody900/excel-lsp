# Changelog

All notable changes to Excel LSP will be documented in this file. The format is
based on Keep a Changelog, and the project intends to follow Semantic
Versioning after its first public release.

## [Unreleased]

No changes yet.

## [0.1.0] - 2026-07-29

### Added

- Project packaging, locked dependencies, CI scaffold, deterministic fixture
  framework, release planning, and Codex-first repository instructions.
- A provenance-checked VBA fixture asset for later macro-preservation tests.
- Streaming, namespace-tolerant OOXML parsing for supported modern Excel
  packages, including normalized cells, tables, names, styles, formulas,
  external-link metadata, and calculation properties.
- A WAL-backed SQLite index with R*Tree or interval fallback, canonical exports,
  full and incremental lifecycle handling, structured errors, and hash-stable
  no-op refreshes.
- Deterministic F01/F07 generation and pinned-openpyxl oracle evidence for the
  completed parser/index foundation.
- Sparse, ListObject-first region and header inference; bounded column
  profiles; frozen public symbol IDs; deterministic F02/F03/F12/F13/F14/F20
  fixtures; a spatial table-barrier index; height-independent merged-range
  spans; bounded 0–8 gap tolerance; component-first, table-aware region
  partitioning with root-local sparse member indexes; and a compact workbook
  map with golden character/token budgets.
- The P2 public-repository skeleton, security scope, contribution guide, and an
  exhaustive README claims-to-artifacts contract. P2's mechanics, test, and
  user-authorized early repository gates all have approving verdicts.
- Formula tokenization and reference classification for A1, names, 3-D,
  structured, external, spill, implicit-intersection, and dynamic-reference
  forms; typed LET/LAMBDA and first-class callable flow; exact `_xlpm.` stored
  local namespaces; conservative lexical colon/intersection handling; modern
  shared-formula translation; exact R1C1 formula blocks; translated,
  context-equivalent structured-reference tiling; clamped edge extrusion;
  F07/F19 semantic goldens; and atomic incremental formula lifecycle persistence.
- Ranked bidirectional dependency mirrors with exact semantic ordering;
  precedents, dependents, bounded traces and shortest paths; R*Tree/interval
  parity; two-stage circular analysis; corruption-resistant live seals and
  capability-isolated SQLite tracking; and approved P4 mechanics/test evidence.
- A complete typed diagnostics catalog; persisted cached-error, external-link,
  and volatile-block findings; deterministic sheet/severity/code filtering;
  four P5 fixtures with exact golden and independent parser-oracle coverage;
  and fail-closed diagnostic-row validation.
- Surgical OOXML cell editing with inline strings, formula-cache invalidation,
  shared-formula expansion, array refusal, calc-chain cleanup, ordered element
  insertion, lock/conflict checks, validated atomic replacement, direct index
  patching, transitive `I_STALE` propagation, A1/R1C1 semantic column fills,
  F16/F21 complete part manifests, a 50-script preservation property, and a
  desktop-Excel recalculation smoke test.
- A verified 14-tool FastMCP stdio server and shared Typer CLI with generated
  schemas, annotations, progress, freshness, generation-bound cursors,
  realpath confinement, guarded regex search, deterministic response caps,
  surgical write integration, and real-client subprocess conformance.
- Three live Excel-authored workbooks, a complete desktop-Excel protocol with
  VBA execution, open-workbook refusal, chart/image preservation, numbered
  screenshots, and a lineage demo GIF.
- Six deterministic benchmark tasks, exact-answer graders, scripted and two-run
  headless-Codex results, raw rows, timing data, and five reproducible charts.
- Clean wheel and `uvx` install probes, Codex MCP registration evidence, fair
  pinned-revision comparisons, public release metadata, CI across three
  operating systems and three Python versions, and registry submission copy.
- Correct MCP implementation metadata so clients see Excel LSP version `0.1.0`
  instead of the installed MCP SDK version.

### Benchmark result

- Excel LSP meets S5 on the disclosed deterministic workbook-payload metric:
  3,410 tool-result tokens versus 222,289 for naive dump (65.2× reduction),
  with 12/12 versus 8/12 exact headless answers. Full CLI usage, including
  fixed agent context and schemas, remains published as a separate secondary
  measurement.

### Planned after v0.1.0

- Rename refactoring for sheets, columns, and defined names.
- A real LSP wire protocol, linked-workbook workspaces, value-level workbook
  diff, watch mode, `.xlsb`, Google Sheets, datetime writes, and stable aliases.

## Release links

[Unreleased]: https://github.com/Brody900/excel-lsp/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/Brody900/excel-lsp/releases/tag/v0.1.0
