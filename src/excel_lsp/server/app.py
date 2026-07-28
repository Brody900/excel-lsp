"""FastMCP stdio server exposing the frozen Excel LSP tool set."""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Callable, Sequence
from typing import Annotated, Any

from mcp.server.fastmcp import Context, FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.server.models import WriteCellInput
from excel_lsp.server.service import ToolService, fit_tool_envelope

INSTRUCTIONS = (
    "Start with open_workbook, inspect a focused region with get_region_schema, then use "
    "trace_precedents, trace_dependents, or trace_path for lineage. Call read_range last and "
    "request the fewest values needed; Excel LSP deliberately avoids bulk workbook dumps. Every "
    "tool checks file freshness. A true reindexed flag means the sidecar was refreshed. Value "
    "responses report stale calculation ranges; recalculate in Excel and call refresh with "
    "recalculated=true when cached results matter. Pagination cursors bind to one index "
    "generation, "
    "so a refresh or write invalidates old cursors. Review destructive writes before invoking them."
)

READ_ANNOTATIONS = ToolAnnotations(readOnlyHint=True, openWorldHint=False)
WRITE_ANNOTATIONS = ToolAnnotations(
    readOnlyHint=False,
    destructiveHint=True,
    openWorldHint=False,
)

mcp = FastMCP("excel-lsp", instructions=INSTRUCTIONS)
service = ToolService()
logger = logging.getLogger(__name__)


async def _invoke(
    method: Callable[..., dict[str, Any]], /, *args: Any, **kwargs: Any
) -> dict[str, Any]:
    try:
        return fit_tool_envelope(await asyncio.to_thread(method, *args, **kwargs))
    except ExcelLSPError as exc:
        return fit_tool_envelope(exc.as_dict())
    except Exception:
        logger.exception("Unexpected Excel LSP tool failure")
        return fit_tool_envelope(
            ExcelLSPError(
                ErrorCode.INTERNAL,
                "Excel LSP encountered an unexpected internal failure.",
                hint=(
                    "Retry once; if it persists, run the equivalent CLI command with debug logging."
                ),
            ).as_dict()
        )


async def _report_sheet_progress(path: str, ctx: Context) -> None:
    try:
        sheets = await asyncio.to_thread(service.sheet_names, path)
    except ExcelLSPError:
        return
    total = len(sheets)
    for number, sheet in enumerate(sheets, 1):
        await ctx.report_progress(number, total, f"Indexed {sheet}")


async def _invoke_with_progress(
    method: Callable[..., dict[str, Any]],
    path: str,
    ctx: Context,
    *args: Any,
) -> dict[str, Any]:
    loop = asyncio.get_running_loop()
    reported = 0

    def report(number: int, total: int, sheet: str) -> None:
        nonlocal reported
        reported += 1
        notification = ctx.report_progress(number, total, f"Indexed {sheet}")
        asyncio.run_coroutine_threadsafe(notification, loop).result()

    result = await _invoke(method, path, *args, progress=report)
    if "error" not in result and reported == 0:
        await _report_sheet_progress(path, ctx)
    return result


@mcp.tool(annotations=READ_ANNOTATIONS)
async def open_workbook(path: str, ctx: Context) -> dict[str, Any]:
    """Open or refresh a workbook and return its compact semantic map."""
    return await _invoke_with_progress(service.open_workbook, path, ctx)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def refresh(path: str, ctx: Context, recalculated: bool = False) -> dict[str, Any]:
    """Refresh a workbook index and optionally clear recalculated staleness."""
    return await _invoke_with_progress(service.refresh, path, ctx, recalculated)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def list_symbols(
    path: str, query: str = "", kinds: list[str] | None = None
) -> dict[str, Any]:
    """List stable semantic symbol ids with short descriptors."""
    return await _invoke(service.list_symbols, path, query, kinds)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def get_region_schema(path: str, region_id: str) -> dict[str, Any]:
    """Describe one semantic region, including bounded samples and validations."""
    return await _invoke(service.get_region_schema, path, region_id)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def read_range(
    path: str,
    ref: str,
    cursor: str | None = None,
    max_cells: Annotated[int, Field(ge=1, le=200)] = 200,
) -> dict[str, Any]:
    """Read one generation-bound page of at most 200 cell values."""
    return await _invoke(service.read_range, path, ref, cursor, max_cells)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def find(
    path: str,
    pattern: str,
    in_: Annotated[
        list[str] | None,
        Field(validation_alias="in", serialization_alias="in_"),
    ] = None,
    sheet: str | None = None,
    max: Annotated[int, Field(ge=1, le=50)] = 50,
) -> dict[str, Any]:
    """Regex-search bounded values, headers, formulas, and names."""
    return await _invoke(service.find, path, pattern, in_, sheet, max)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def trace_precedents(
    path: str,
    ref_or_symbol: str,
    depth: Annotated[int, Field(ge=0, le=8)] = 2,
    max_nodes: Annotated[int, Field(ge=1, le=200)] = 200,
) -> dict[str, Any]:
    """Trace a bounded upstream dependency tree."""
    return await _invoke(service.trace_precedents, path, ref_or_symbol, depth, max_nodes)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def trace_dependents(
    path: str,
    ref_or_symbol: str,
    depth: Annotated[int, Field(ge=0, le=8)] = 2,
    max_nodes: Annotated[int, Field(ge=1, le=200)] = 200,
) -> dict[str, Any]:
    """Trace a bounded downstream dependency tree."""
    return await _invoke(service.trace_dependents, path, ref_or_symbol, depth, max_nodes)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def trace_path(
    path: str,
    from_ref_or_symbol: str,
    to_ref_or_symbol: str,
    max_paths: Annotated[int, Field(ge=1, le=50)] = 3,
    max_depth: Annotated[int, Field(ge=0, le=12)] = 12,
) -> dict[str, Any]:
    """Find bounded shortest dependency paths between two areas or symbols."""
    return await _invoke(
        service.trace_path,
        path,
        from_ref_or_symbol,
        to_ref_or_symbol,
        max_paths,
        max_depth,
    )


@mcp.tool(annotations=READ_ANNOTATIONS)
async def explain_formula(path: str, ref: str) -> dict[str, Any]:
    """Explain one formula cell in A1, R1C1, block, and diagnostic terms."""
    return await _invoke(service.explain_formula, path, ref)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def get_diagnostics(
    path: str,
    sheet: str | None = None,
    severity: str | None = None,
    code: str | None = None,
    max: Annotated[int, Field(ge=1, le=100)] = 100,
) -> dict[str, Any]:
    """Return filtered workbook diagnostics with aggregate counts."""
    return await _invoke(service.get_diagnostics, path, sheet, severity, code, max)


@mcp.tool(annotations=READ_ANNOTATIONS)
async def profile_column(path: str, col_symbol_or_ref: str) -> dict[str, Any]:
    """Return bounded numeric or categorical statistics for one column."""
    return await _invoke(service.profile_column, path, col_symbol_or_ref)


@mcp.tool(annotations=WRITE_ANNOTATIONS)
async def write_cells(path: str, cells: Sequence[WriteCellInput]) -> dict[str, Any]:
    """Surgically write at most 500 qualified cells."""
    return await _invoke(service.write_cells, path, cells)


@mcp.tool(annotations=WRITE_ANNOTATIONS)
async def set_column_formula(
    path: str,
    col_symbol: str,
    formula: str,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Fill one semantic column with a translated A1 or R1C1 formula."""
    return await _invoke(service.set_column_formula, path, col_symbol, formula, overwrite)


def run() -> None:
    """Run the production stdio transport."""
    mcp.run(transport="stdio")


__all__ = ["INSTRUCTIONS", "mcp", "run"]
