"""Tests for the Phase 0 CLI surface."""

from typer.testing import CliRunner

from excel_lsp import __version__
from excel_lsp.cli import app

runner = CliRunner()


def test_help_describes_the_cli() -> None:
    result = runner.invoke(app, ["--help"])

    assert result.exit_code == 0
    assert "Excel LSP command-line interface." in result.stdout
    assert "--version" in result.stdout


def test_version_reports_package_version() -> None:
    result = runner.invoke(app, ["--version"])

    assert result.exit_code == 0
    assert result.stdout.strip() == __version__
