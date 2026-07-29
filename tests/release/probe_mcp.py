"""Probe an installed Excel LSP executable through a real MCP stdio client."""

from __future__ import annotations

import argparse
import asyncio
import json
import os
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

EXPECTED_TOOLS = {
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
}


async def _probe(
    executable: Path,
    workbook: Path,
    uvx_package: str | None,
    expected_version: str,
) -> dict[str, object]:
    environment = os.environ.copy()
    environment["EXCEL_LSP_ROOT"] = str(workbook.parent)
    parameters = StdioServerParameters(
        command=str(executable),
        args=["--from", uvx_package, "excel-lsp", "serve"] if uvx_package else ["serve"],
        cwd=workbook.parent,
        env=environment,
    )
    async with stdio_client(parameters) as (read_stream, write_stream):  # noqa: SIM117
        async with ClientSession(read_stream, write_stream) as session:
            initialized = await session.initialize()
            tools = await session.list_tools()
            names = {tool.name for tool in tools.tools}
            assert names == EXPECTED_TOOLS
            opened = await session.call_tool("open_workbook", {"path": str(workbook)})
            assert not opened.isError
            assert initialized.instructions
            assert initialized.serverInfo.name == "excel-lsp"
            assert initialized.serverInfo.version == expected_version
            return {
                "implementation": initialized.serverInfo.name,
                "version": initialized.serverInfo.version,
                "toolCount": len(names),
                "instructions": True,
                "openWorkbook": True,
            }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("executable", type=Path)
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--uvx-package")
    parser.add_argument("--expected-version", default="0.1.0")
    args = parser.parse_args()
    result = asyncio.run(
        _probe(
            args.executable.resolve(),
            args.workbook.resolve(),
            args.uvx_package,
            args.expected_version,
        )
    )
    print(json.dumps(result, sort_keys=True))


if __name__ == "__main__":
    main()
