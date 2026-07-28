"""Command-line debugging surface for the MCP tool contracts."""

from __future__ import annotations

import json
from importlib import import_module
from pathlib import Path
from typing import Annotated, Any

import typer

from excel_lsp import __version__
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.server.service import ToolService, fit_tool_envelope

app = typer.Typer(
    add_completion=False,
    help="Excel LSP command-line interface. Semantic navigation and editing for workbooks.",
    no_args_is_help=True,
    pretty_exceptions_show_locals=False,
)
service = ToolService()


@app.callback(invoke_without_command=True)
def main(
    version: Annotated[
        bool,
        typer.Option("--version", help="Show the installed version and exit.", is_eager=True),
    ] = False,
) -> None:
    """Inspect workbooks through the same bounded contracts used by MCP clients."""
    if version:
        typer.echo(__version__)
        raise typer.Exit()


@app.command()
def serve() -> None:
    """Run the MCP server over stdio."""
    from excel_lsp.server import run

    run()


@app.command("map")
def workbook_map(file: Path) -> None:
    """Print the compact workbook map."""
    _emit_call(service.open_workbook, str(file))


@app.command()
def trace(
    file: Path,
    ref: str,
    deps: Annotated[bool, typer.Option("--deps", help="Trace dependents.")] = False,
    precs: Annotated[bool, typer.Option("--precs", help="Trace precedents.")] = False,
    depth: Annotated[int, typer.Option(min=0, max=8)] = 2,
) -> None:
    """Trace precedents or dependents as JSON."""
    direction = _direction(deps, precs)
    method = service.trace_dependents if direction == "dependents" else service.trace_precedents
    _emit_call(method, str(file), ref, depth, 200)


@app.command("path")
def dependency_path(file: Path, from_ref: str, to_ref: str) -> None:
    """Print shortest dependency paths between two references or symbols."""
    _emit_call(service.trace_path, str(file), from_ref, to_ref, 3, 12)


@app.command("diag")
def diagnostics(file: Path) -> None:
    """Print workbook diagnostics."""
    _emit_call(service.get_diagnostics, str(file))


@app.command("find")
def find_values(file: Path, pattern: str) -> None:
    """Search values, headers, formulas, and names with a guarded regex."""
    _emit_call(service.find, str(file), pattern)


@app.command("schema")
def region_schema(file: Path, region: str) -> None:
    """Print a semantic region schema."""
    _emit_call(service.get_region_schema, str(file), region)


@app.command()
def graph(
    file: Path,
    ref: str,
    deps: Annotated[bool, typer.Option("--deps", help="Trace dependents.")] = False,
    precs: Annotated[bool, typer.Option("--precs", help="Trace precedents.")] = False,
    depth: Annotated[int, typer.Option(min=0, max=8)] = 2,
    mermaid: Annotated[bool, typer.Option("--mermaid", help="Emit a Mermaid flowchart.")] = False,
) -> None:
    """Render a dependency trace as JSON or Mermaid."""
    direction = _direction(deps, precs)
    method = service.trace_dependents if direction == "dependents" else service.trace_precedents
    result = _call(method, str(file), ref, depth, 200)
    if mermaid and "error" not in result:
        typer.echo(_mermaid(result))
        return
    _emit(result)


@app.command()
def bench() -> None:
    """Run the benchmark harness (installed in Phase 8)."""
    try:
        run_cli = import_module("benchmarks.runner").run_cli
    except ImportError:
        _emit(
            ExcelLSPError(
                ErrorCode.NOT_FOUND,
                "The Phase 8 benchmark harness is not installed yet.",
                hint="Complete P8 or run the MCP/CLI conformance suite for P7 evidence.",
            ).as_dict()
        )
        raise typer.Exit(code=1) from None
    raise typer.Exit(code=run_cli())


def _direction(deps: bool, precs: bool) -> str:
    if deps == precs:
        _emit(
            ExcelLSPError(
                ErrorCode.INVALID_VALUE,
                "Choose exactly one of --deps or --precs.",
            ).as_dict()
        )
        raise typer.Exit(code=2)
    return "dependents" if deps else "precedents"


def _call(method: Any, *args: Any) -> dict[str, Any]:
    try:
        return method(*args)
    except ExcelLSPError as exc:
        return exc.as_dict()
    except Exception:
        return ExcelLSPError(
            ErrorCode.INTERNAL,
            "Excel LSP encountered an unexpected internal failure.",
        ).as_dict()


def _emit_call(method: Any, *args: Any) -> None:
    result = _call(method, *args)
    _emit(result)
    if "error" in result:
        raise typer.Exit(code=1)


def _emit(payload: dict[str, Any]) -> None:
    typer.echo(json.dumps(fit_tool_envelope(payload), ensure_ascii=False, indent=2, default=str))


def _mermaid(payload: dict[str, Any]) -> str:
    lines = ["flowchart LR"]
    counter = 0

    def visit(node: dict[str, Any], parent: str | None = None, via: str | None = None) -> None:
        nonlocal counter
        node_id = f"n{counter}"
        counter += 1
        label = str(node.get("symbol") or node.get("ref") or node.get("kind") or "node")
        label = label.replace('"', "'").replace("\n", " ")
        lines.append(f'  {node_id}["{label}"]')
        if parent is not None:
            edge = "" if via is None else f'|"{via.replace(chr(34), chr(39))}"|'
            lines.append(f"  {parent} -->{edge} {node_id}")
        for child in node.get("children", []):
            visit(child, node_id, child.get("via"))

    visit(payload["tree"])
    return "\n".join(lines)


__all__ = ["app"]
