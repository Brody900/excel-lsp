# Changelog

All notable changes to Excel LSP will be documented in this file. The format is
based on Keep a Changelog, and the project intends to follow Semantic
Versioning after its first public release.

## [Unreleased]

Excel LSP is pre-release. No PyPI distribution, GitHub release, or supported
public API is being claimed yet.

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

### Planned before v0.1.0

- The 14-tool stdio MCP server, CLI, response-cap conformance, and Codex
  quickstart verification.
- Live Microsoft Excel evidence, deterministic and headless-Codex benchmarks,
  final documentation, public repository publication, and registry metadata.

## Release links

Release comparison links will be added when the public GitHub repository and
the first tag exist. Leaving them absent is intentional; placeholder URLs would
be misleading.
