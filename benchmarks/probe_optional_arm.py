"""List the optional comparison server's MCP tools for compatibility triage."""

from __future__ import annotations

import asyncio
import json

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


async def probe() -> list[dict[str, object]]:
    parameters = StdioServerParameters(command="uvx", args=["excel-mcp-server", "stdio"])
    async with (
        stdio_client(parameters) as (reader, writer),
        ClientSession(reader, writer) as session,
    ):
        await session.initialize()
        result = await session.list_tools()
        return [
            {
                "name": tool.name,
                "required": tool.inputSchema.get("required", []),
                "properties": sorted(tool.inputSchema.get("properties", {})),
            }
            for tool in result.tools
        ]


if __name__ == "__main__":
    print(json.dumps(asyncio.run(probe()), indent=2))
