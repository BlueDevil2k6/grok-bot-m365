"""Spawn the server over stdio exactly as an MCP host would, list tools, call auth_status."""
import asyncio, os
from mcp import ClientSession
from mcp.client.stdio import stdio_client, StdioServerParameters

async def main():
    p = StdioServerParameters(command="uv", args=["run", "--directory", os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "python", "server.py"],
                              env={**os.environ, "M365_CLIENT_ID": "00000000-0000-0000-0000-000000000000",
                                   "M365_TOKEN_CACHE": "/tmp/grok-bot-m365-smoke-cache.json"})
    async with stdio_client(p) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()
            tools = await s.list_tools()
            print("TOOLS:", sorted(t.name for t in tools.tools))
            res = await s.call_tool("list_aliases", {})
            print("list_aliases ->", res.is_error, res.content[0].text[:120])
asyncio.run(main())
