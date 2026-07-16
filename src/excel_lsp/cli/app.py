"""Minimal Phase 0 command-line interface."""

from typing import Annotated

import typer

from excel_lsp import __version__

app = typer.Typer(
    add_completion=False,
    help="Excel LSP command-line interface.",
    no_args_is_help=True,
    pretty_exceptions_show_locals=False,
)


@app.callback(invoke_without_command=True)
def main(
    version: Annotated[
        bool,
        typer.Option("--version", help="Show the installed version and exit.", is_eager=True),
    ] = False,
) -> None:
    """Show package information; workbook commands arrive in later phases."""
    if version:
        typer.echo(__version__)
        raise typer.Exit()
