"""Repository-contract tests for the P2 public README skeleton."""

from __future__ import annotations

import json
import re
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
README = ROOT / "README.md"
CLAIMS = ROOT / "docs" / "evidence" / "readme-claims-to-artifacts.md"
TOOL_REFERENCE = ROOT / "docs" / "tool-reference.md"
RAW_RESULTS_INDEX = ROOT / "benchmarks" / "results" / "README.md"
PLAN = ROOT / "PLAN.md"
ARCHITECTURE = ROOT / "docs" / "architecture.md"

ONE_LINER = (
    "An LSP for Excel: semantic index + MCP server so AI agents navigate workbooks by "
    "symbols, references, and diagnostics — not by reading 50,000 rows."
)
QUALIFIER = (
    "*(LSP-style: the ideas — symbols, references, diagnostics, incremental index — not "
    "the LSP wire protocol.)*"
)
TRADEMARK_FOOTER = "Not affiliated with Microsoft. Excel is a trademark of Microsoft Corporation."

README_SECTIONS = (
    "## 60-second lineage demo",
    "## Quickstart",
    "## Tools",
    "## Architecture",
    "## Benchmarks",
    "## Comparison",
    "## How it works",
    "## Security & scope",
    "## Limitations and roadmap",
    "## Evidence",
)

TOOLS = (
    "open_workbook",
    "refresh",
    "list_symbols",
    "get_region_schema",
    "read_range",
    "find",
    "trace_precedents",
    "trace_dependents",
    "trace_path",
    "explain_formula",
    "get_diagnostics",
    "profile_column",
    "write_cells",
    "set_column_formula",
)


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _markdown_anchors(path: Path) -> set[str]:
    anchors: set[str] = set()
    for line in _read(path).splitlines():
        if not line.startswith("#"):
            continue
        heading = line.lstrip("#").strip().casefold()
        heading = re.sub(r"[^\w\s-]", "", heading)
        anchors.add(re.sub(r"[\s-]+", "-", heading).strip("-"))
    return anchors


def _assert_exact_reference(reference: str) -> None:
    if "::" in reference:
        path_text, node_id = reference.split("::", maxsplit=1)
        path = ROOT / path_text
        assert path.exists(), reference
        function_name = node_id.split("[", maxsplit=1)[0]
        assert f"def {function_name}(" in _read(path), reference
        return

    path_text, separator, anchor = reference.partition("#")
    path = ROOT / path_text
    assert path.exists(), reference
    if separator:
        assert anchor in _markdown_anchors(path), reference


def test_readme_has_frozen_positioning_and_section_order() -> None:
    readme = _read(README)

    assert readme.startswith("# Excel LSP\n\n")
    assert f"{ONE_LINER}\n\n{QUALIFIER}" in readme
    assert "Benchmark hero chart — planned for P8" in readme
    assert "P8 HERO SLOT" in readme

    positions = [readme.index(heading) for heading in README_SECTIONS]
    assert positions == sorted(positions)
    assert (
        readme.index(QUALIFIER)
        < readme.index("Benchmark hero chart — planned for P8")
        < readme.index("P8 HERO SLOT")
        < positions[0]
    )
    assert readme.rstrip().endswith(TRADEMARK_FOOTER)


def test_readme_lists_exactly_the_frozen_fourteen_tools() -> None:
    readme = _read(README)
    tools_section = readme.split("## Tools", maxsplit=1)[1].split("## Architecture", maxsplit=1)[0]
    table_tools = tuple(re.findall(r"^\| `([^`]+)` \|", tools_section, flags=re.MULTILINE))

    assert table_tools == TOOLS
    assert "The first 12 tools are read tools." in tools_section
    assert "The final two are destructive write tools" in tools_section

    tool_headings = tuple(re.findall(r"^## `([^`]+)`$", _read(TOOL_REFERENCE), flags=re.MULTILINE))
    assert tool_headings == TOOLS


def test_readme_is_codex_first_and_labels_future_evidence() -> None:
    readme = _read(README)
    lowered = readme.casefold()

    assert "codex mcp add excel-lsp -- uvx excel-lsp serve" in readme
    assert "[mcp_servers.excel-lsp]" in readme
    assert not re.search(r"\bclaude\s+(?:mcp|exec|-p)\b", lowered)
    assert "This is a generic MCP-client example, not Codex's native" in readme

    required_status_labels = (
        "Development status",
        "Planned for P8",
        "Planned v0.1.0 quickstart",
        "Planned for P9",
        "currently makes no measured performance",
    )
    for label in required_status_labels:
        assert label in readme


def test_readme_preserves_security_scope_and_limitations() -> None:
    readme = _read(README)
    normalized_readme = " ".join(readme.split())

    required_text = (
        "Pre-release security boundary; P7/P9 verification remains pending",
        "implemented core operates on local workbook paths and makes no runtime",
        "network requests. MCP path confinement",
        "default will be unrestricted local-path access",
        "every OOXML part not deliberately modified stays byte-identical",
        "P6 core verified; P8 live protocol pending",
        "does not recalculate formulas",
        "Verified P3/P5",
        "are flaggable but opaque to static dependency analysis",
        "every inferred region exposes a confidence score",
        "P6 candidate",
        "Written strings use OOXML inline strings",
        "P6 core implemented; P7 exposure pending",
        "Datetime cell writes are rejected in v0.1.0",
        "Writes inside multi-cell array formulas are refused",
        "Verified P3",
        "Dynamic-array spill extents are not statically tracked",
        "Flagship v1.x item",
        "powered by the dependency graph",
        "A real LSP wire-protocol server",
        TRADEMARK_FOOTER,
    )
    for text in required_text:
        assert " ".join(text.split()) in normalized_readme


def test_benchmark_section_links_the_planned_raw_results_index() -> None:
    readme = _read(README)
    raw_results = _read(RAW_RESULTS_INDEX)
    benchmark_section = readme.split("## Benchmarks", maxsplit=1)[1].split(
        "## Comparison", maxsplit=1
    )[0]
    normalized_section = " ".join(benchmark_section.split())

    assert "[raw results index](benchmarks/results/README.md)" in benchmark_section
    assert "placeholder until P8 commits measured rows" in normalized_section
    for filename in (
        "environment.json",
        "scripted.csv",
        "llm-eval.jsonl",
        "accuracy.csv",
        "index-timing.csv",
        "audit-cost.json",
    ):
        assert f"`{filename}`" in raw_results
    assert "`benchmarks/check.py`" in raw_results
    assert "`excel-lsp bench`" in raw_results
    assert "Benchmarks show \u2265 10\u00d7 token reduction vs." in raw_results


def test_readme_local_links_exist() -> None:
    readme = _read(README)
    targets = re.findall(r"\[[^]]+\]\(([^)]+)\)", readme)

    assert targets
    for target in targets:
        if target.startswith(("http://", "https://", "#")):
            continue
        path_text = target.split("#", maxsplit=1)[0]
        assert (ROOT / path_text).exists(), target


def test_codex_and_generic_mcp_examples_are_equivalent() -> None:
    codex_config = tomllib.loads(_read(ROOT / "examples" / "codex.config.toml"))
    generic_config = json.loads(_read(ROOT / "examples" / "mcp.json"))

    codex_server = codex_config["mcp_servers"]["excel-lsp"]
    generic_server = generic_config["mcpServers"]["excel-lsp"]
    expected = {"command": "uvx", "args": ["excel-lsp", "serve"]}

    assert codex_server == expected
    assert generic_server == expected


def test_claims_matrix_has_unique_mapped_rows_and_statuses() -> None:
    claims = _read(CLAIMS)
    lines = [line for line in claims.splitlines() if line.startswith("|")]
    assert lines[0] == (
        "| ID | README item or claim | Phase | Status | Required committed artifact |"
    )

    rows: dict[str, tuple[str, str, str, str]] = {}
    for line in lines[2:]:
        cells = tuple(cell.strip() for cell in line.strip("|").split("|"))
        assert len(cells) == 5, line
        claim_id, claim, phase, status, artifact = cells
        assert claim_id not in rows
        assert claim and phase and status and artifact
        assert "TBD" not in artifact.upper()
        assert re.fullmatch(r"P\d(?:/P\d)*", phase)
        assert re.fullmatch(
            r"(?:Verified P\d|Planned P\d|Scope declaration)",
            status,
        )
        references = re.findall(r"`([^`]+)`", artifact)
        assert references, line
        residual = re.sub(r"`[^`]+`", "", artifact).replace("<br>", "").strip()
        assert not residual, line
        rows[claim_id] = (claim, phase, status, artifact)

    required_ids = {
        "P1-FOUND",
        "P2-FOUND",
        "POS-01",
        "POS-02",
        "POS-03",
        "S1",
        "S2",
        "S2-CAP",
        "S3",
        "S4",
        "S5",
        "S6",
        "S7",
        "HERO-01",
        "DEMO-01",
        "QS-01",
        "QS-02",
        "QS-03",
        "QS-04",
        "TOOLS-15",
        "ARCH-01",
        "ARCH-02",
        "BENCH-01",
        "BENCH-02",
        "BENCH-CLI",
        "BENCH-03",
        "BENCH-04",
        "COMP-01",
        "COMP-02",
        "COMP-03",
        "COMP-04",
        "COMP-05",
        "COMP-06",
        "HOW-01",
        "HOW-02",
        "HOW-03",
        "SEC-01",
        "SEC-02",
        "SEC-03",
        "LIM-01",
        "LIM-02",
        "LIM-03",
        "LIM-04",
        "LIM-05",
        "LIM-06",
        "LIM-07",
        "NGOAL-01",
        "ROAD-01",
        "ROAD-02",
        "EVID-01",
        "LEGAL-01",
    }
    required_ids.update(f"TOOL-{number:02d}" for number in range(1, 15))

    assert set(rows) == required_ids
    assert rows["S5"][0] == (
        "Benchmarks show \u2265 10\u00d7 token reduction vs. the naive-dump baseline on "
        "the defined task suite, with equal-or-better task accuracy in LLM evals."
    )
    assert (
        "`tests/unit/test_cli.py::test_bench_command_runs_reproducible_harness`"
        in rows["BENCH-CLI"][3]
    )


def test_current_claim_artifact_paths_nodes_and_anchors_resolve() -> None:
    claims = _read(CLAIMS)
    for line in claims.splitlines():
        if not line.startswith("|") or line.startswith(("| ID ", "|---")):
            continue
        cells = tuple(cell.strip() for cell in line.strip("|").split("|"))
        status = cells[3]
        if status not in {
            "Verified P1",
            "Verified P2",
            "Verified P3",
            "Verified P4",
            "Verified P5",
        }:
            continue
        for reference in re.findall(r"`([^`]+)`", cells[4]):
            _assert_exact_reference(reference)


def test_readme_links_current_positive_comparison_claims_to_p1_evidence() -> None:
    readme = _read(README)

    assert "[P1 evidence available](docs/evidence/p1-foundation.md#delivered-contracts)" in readme
    assert "[P1 evidence available](docs/evidence/p1-foundation.md#invariant-evidence)" in readme


def test_repository_skeleton_documents_are_substantive_and_pre_release() -> None:
    for filename in ("SECURITY.md", "CONTRIBUTING.md", "CHANGELOG.md"):
        content = _read(ROOT / filename)
        assert len(content.splitlines()) >= 25

    security = _read(ROOT / "SECURITY.md")
    changelog = _read(ROOT / "CHANGELOG.md")
    assert "No public version is supported yet" in security
    assert "Excel LSP is pre-release" in changelog


def test_review_governance_current_state_is_consistent() -> None:
    plan = _read(PLAN)
    claims = _read(CLAIMS)
    architecture = _read(ARCHITECTURE)
    current_state_paths = (
        PLAN,
        CLAIMS,
        ARCHITECTURE,
        ROOT / "docs" / "evidence" / "README.md",
        ROOT / "docs" / "evidence" / "p2-regions-map.md",
        ROOT / "docs" / "index-internals.md",
        ROOT / "CHANGELOG.md",
    )

    assert "Total formal verdicts recorded: 64." in plan
    assert "0 nominal slots remain" in plan
    assert "The minimum completion path is now 30 total invocations" in plan
    assert "leaving no pooled\ncontingency slots" in plan
    assert "fresh re-review after any subsequent `REVISE`" in plan
    assert "2026-07-27 unbounded-review amendment" in plan
    assert "supersedes the first amendment's 30-invocation hard stop" in plan
    assert "2026-07-28 single-reviewer amendment" in plan
    assert "exactly one independent reviewer evaluates\nboth domains" in plan
    assert "The first required early P2 R-repo invocation is charged as `REVISE`." in claims
    assert "a fresh stateless R-repo re-review\nfollows each `REVISE`" in claims
    assert "The early P2 repository gate\nis approved." in claims
    assert "Its R-mech, R-test, and early R-repo gates have approved" in architecture
    assert (
        "| P2 | Regions, headers, stable symbols, compact workbook map | Verified |" in architecture
    )
    assert "Until its R-mech and R-test gates approve" not in architecture
    assert "all three required review\napprovals" not in architecture

    for path in current_state_paths:
        content = _read(path)
        assert "exactly one R-repo" not in content, path
        assert "one-shot" not in content, path
        assert "governance decision for any required" not in content, path
