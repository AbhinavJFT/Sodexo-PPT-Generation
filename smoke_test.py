"""Quick test that the MCP server starts and we can list + call tools."""
import asyncio
import sys
from pathlib import Path

BASE = Path(__file__).parent
MCP_BIN = str((BASE / ".venv" / "bin" / "ppt_mcp_server").resolve())

sys.path.insert(0, str(BASE))
from mcp_client import mcp_session, mcp_tools_to_groq_schemas, parse_tool_result, EXPOSED_TOOLS


async def main():
    async with mcp_session(MCP_BIN) as s:
        tr = await s.list_tools()
        names = [t.name for t in tr.tools]
        print(f"Total tools from server: {len(names)}")
        exposed = [n for n in names if n in EXPOSED_TOOLS]
        print(f"Exposed to LLM: {len(exposed)}")
        missing = EXPOSED_TOOLS - set(names)
        if missing:
            print(f"WARN: exposed but not present: {missing}")

        schemas = mcp_tools_to_groq_schemas(tr.tools)
        print(f"Groq schemas: {len(schemas)}")

        # sanity: call get_server_info
        r = await s.call_tool("get_server_info", {})
        print("get_server_info ->", parse_tool_result(r)[:200])


if __name__ == "__main__":
    asyncio.run(main())
