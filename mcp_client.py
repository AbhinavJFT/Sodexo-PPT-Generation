"""MCP stdio client for Office-PowerPoint-MCP-Server."""
from contextlib import asynccontextmanager
from typing import Any, Dict, List

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

EXPOSED_TOOLS = {
    "open_presentation",
    "get_presentation_info",
    "get_template_file_info",
    "manage_slide_masters",
    "extract_presentation_text",
    "extract_slide_text",
    "get_slide_info",
    "create_presentation_from_template",
    "add_slide",
    "populate_placeholder",
    "add_bullet_points",
    "manage_text",
    "add_table",
    "add_chart",
    "add_shape",
    "add_connector",
    "manage_image",
    "apply_picture_effects",
    "save_presentation",
    "list_presentations",
    "switch_presentation",
    "get_server_info",
}


@asynccontextmanager
async def mcp_session(server_bin: str):
    params = StdioServerParameters(command=server_bin, args=[])
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            yield session


def mcp_tools_to_groq_schemas(tools) -> List[Dict[str, Any]]:
    out = []
    for t in tools:
        if t.name not in EXPOSED_TOOLS:
            continue
        schema = t.inputSchema or {"type": "object", "properties": {}}
        if "type" not in schema:
            schema["type"] = "object"
        out.append({
            "type": "function",
            "function": {
                "name": t.name,
                "description": (t.description or "")[:1024],
                "parameters": schema,
            },
        })
    return out


def parse_tool_result(result) -> str:
    if getattr(result, "isError", False):
        prefix = "[tool_error] "
    else:
        prefix = ""
    if not result.content:
        return prefix + "(no result)"
    parts = []
    for c in result.content:
        parts.append(getattr(c, "text", None) or str(c))
    return prefix + "\n".join(parts)
