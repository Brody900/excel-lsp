"""Tests for the P7 CLI debugging surface."""

import json
from pathlib import Path

from typer.testing import CliRunner

from excel_lsp import __version__
from excel_lsp.cli import app

runner = CliRunner()
FIXTURE = Path(__file__).parents[1] / "fixtures" / "generated" / "cross_sheet_model.xlsx"


def test_help_describes_the_cli() -> None:
    result = runner.invoke(app, ["--help"])

    assert result.exit_code == 0
    assert "Excel LSP command-line interface." in result.stdout
    for command in ("serve", "map", "trace", "path", "diag", "find", "schema", "graph", "bench"):
        assert command in result.stdout


def test_version_reports_package_version() -> None:
    result = runner.invoke(app, ["--version"])

    assert result.exit_code == 0
    assert result.stdout.strip() == __version__


def test_map_and_read_commands_emit_tool_json() -> None:
    mapped = runner.invoke(app, ["map", str(FIXTURE)])
    assert mapped.exit_code == 0
    assert json.loads(mapped.stdout)["workbook"] == FIXTURE.name

    cases = (
        ["trace", str(FIXTURE), "Inputs!B2", "--deps", "--depth", "1"],
        ["path", str(FIXTURE), "Inputs!B2", "Summary!C2"],
        ["diag", str(FIXTURE)],
        ["find", str(FIXTURE), "Revenue"],
        ["schema", str(FIXTURE), "region:Inputs:0"],
    )
    for arguments in cases:
        result = runner.invoke(app, arguments)
        assert result.exit_code == 0, result.stdout
        assert "error" not in json.loads(result.stdout)


def test_graph_mermaid_and_direction_validation() -> None:
    rendered = runner.invoke(
        app,
        ["graph", str(FIXTURE), "Inputs!B2", "--deps", "--mermaid", "--depth", "1"],
    )
    assert rendered.exit_code == 0
    assert rendered.stdout.startswith("flowchart LR")
    assert "cell:Inputs!B2" in rendered.stdout

    invalid = runner.invoke(app, ["trace", str(FIXTURE), "Inputs!B2"])
    assert invalid.exit_code == 2
    assert json.loads(invalid.stdout)["error"]["code"] == "E_INVALID_VALUE"


def test_bench_command_runs_reproducible_harness(tmp_path: Path, monkeypatch) -> None:
    from benchmarks import run_scripted

    output = tmp_path / "scripted.csv"
    monkeypatch.setattr(run_scripted, "DEFAULT_OUTPUT", output)
    from benchmarks import runner as benchmark_runner

    monkeypatch.setattr(benchmark_runner, "DEFAULT_OUTPUT", output)

    result = runner.invoke(app, ["bench"])

    assert result.exit_code == 0, result.stdout
    payload = json.loads(result.stdout)
    assert payload == {
        "mode": "scripted",
        "rows": 12,
        "failed": 0,
        "output": str(output),
        "next": "Run benchmarks/run_llm_eval.py for the cost-guarded headless mode.",
    }
    assert output.exists()
