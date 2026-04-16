"""Groq tool-calling loop that drives the MCP PowerPoint server."""
import json
import logging
from typing import Any, Awaitable, Callable, Dict, List

log = logging.getLogger("agent")

SYSTEM_PROMPT = """You are a PowerPoint presentation generator using the Office-PowerPoint-MCP-Server tools.

Reference deck (read-only): {template_path}
Output path (MUST save here): {output_path}

Your task: create a NEW presentation that preserves the visual FORMAT of the reference deck (theme, fonts, colors, slide masters, layouts) but contains NEW content based on the user's request.

Required workflow:
1. get_template_file_info(template_path="{template_path}") -- learn layouts/dimensions.
2. open_presentation(file_path="{template_path}", id="src") -- load reference for inspection.
3. manage_slide_masters(operation="list", presentation_id="src")
4. manage_slide_masters(operation="get_layouts", master_index=0, presentation_id="src") -- note layout names and indices.
5. (Optional) extract_presentation_text(presentation_id="src") -- see existing content style.
6. create_presentation_from_template(template_path="{template_path}", id="out") -- clones masters/theme.
7. For each new slide, call add_slide(layout_index=<suitable>, title=<str>, presentation_id="out"). Then use populate_placeholder or add_bullet_points to fill in the body. If you need placeholder indices, call get_slide_info first.
8. save_presentation(file_path="{output_path}", presentation_id="out") -- final step.

Strict rules:
- Always pass presentation_id explicitly ("src" for reading, "out" for writing).
- Pick layout_index from the reference's slide_layouts -- prefer "Title Slide" for the cover and "Title and Content" / "Content" style layouts for body slides.
- NEVER call auto_generate_presentation, apply_slide_template, create_slide_from_template, or create_presentation_from_templates -- those override the reference theme.
- After a successful save_presentation call, reply with a brief 1-2 sentence confirmation and STOP calling tools.
- Keep the new deck focused: 4-8 slides unless the user asks otherwise.
"""


async def run_agent(
    groq_client,
    model: str,
    user_prompt: str,
    template_path: str,
    output_path: str,
    tool_schemas: List[Dict[str, Any]],
    call_mcp_tool: Callable[[str, Dict[str, Any]], Awaitable[str]],
    max_iters: int = 40,
) -> Dict[str, Any]:
    messages: List[Dict[str, Any]] = [
        {"role": "system", "content": SYSTEM_PROMPT.format(
            template_path=template_path, output_path=output_path)},
        {"role": "user", "content": user_prompt},
    ]
    trace = []
    saved = False

    for it in range(max_iters):
        resp = groq_client.chat.completions.create(
            model=model,
            messages=messages,
            tools=tool_schemas,
            tool_choice="auto",
            temperature=0.2,
            max_tokens=2048,
        )
        msg = resp.choices[0].message
        tool_calls = msg.tool_calls or []

        assistant_entry: Dict[str, Any] = {"role": "assistant", "content": msg.content or ""}
        if tool_calls:
            assistant_entry["tool_calls"] = [
                {
                    "id": tc.id,
                    "type": "function",
                    "function": {"name": tc.function.name, "arguments": tc.function.arguments or "{}"},
                }
                for tc in tool_calls
            ]
        messages.append(assistant_entry)

        if not tool_calls:
            return {
                "saved": saved,
                "iterations": it + 1,
                "trace": trace,
                "final_message": msg.content or "",
            }

        for tc in tool_calls:
            name = tc.function.name
            try:
                args = json.loads(tc.function.arguments or "{}")
            except json.JSONDecodeError as e:
                args = {}
                err = f"Invalid JSON arguments: {e}"
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": err})
                trace.append({"tool": name, "args": tc.function.arguments, "result": err})
                continue

            log.info("tool %s args=%s", name, args)
            try:
                result_text = await call_mcp_tool(name, args)
            except Exception as e:
                result_text = json.dumps({"error": f"{type(e).__name__}: {e}"})

            trace.append({"tool": name, "args": args, "result": result_text[:400]})
            if name == "save_presentation" and not result_text.startswith("[tool_error]") and '"error"' not in result_text:
                saved = True

            messages.append({
                "role": "tool",
                "tool_call_id": tc.id,
                "content": result_text[:6000],
            })

    return {
        "saved": saved,
        "iterations": max_iters,
        "trace": trace,
        "final_message": "max iterations reached",
    }
