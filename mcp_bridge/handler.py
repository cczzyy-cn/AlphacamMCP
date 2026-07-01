"""
MCP tool dispatcher — routes tool calls to registered handlers.
"""

from __future__ import annotations

import json
import logging
from typing import Any

from mcp.types import TextContent, CallToolResult

from alphacam_com import AlphaCAM, AlphaCAMError, AlphaCAMNotRunning
from .errors import ToolError
from .tools import get_tool_func
from .config import get_prog_id, get_visible
from . import docs as doc_tools

log = logging.getLogger("alphacam-bridge.handler")

# ---------------------------------------------------------------------------
# Global AlphaCAM instance
# ---------------------------------------------------------------------------

_acam: AlphaCAM | None = None


def get_acam() -> AlphaCAM:
    """Get or create the shared AlphaCAM connection."""
    global _acam
    if _acam is None:
        _acam = AlphaCAM(prog_id=get_prog_id(), visible=get_visible())
    return _acam


def _get_acam_instance() -> AlphaCAM | None:
    """Return current acam instance without creating one (for config detection)."""
    return _acam


def set_prog_id(prog_id: str):
    """Override the ProgID (called from CLI --progid)."""
    global _acam
    _acam = None
    import os
    os.environ["ALPHACAM_PROG_ID"] = prog_id


# ---------------------------------------------------------------------------
# Result helpers
# ---------------------------------------------------------------------------

def _json(result: Any) -> list[TextContent]:
    """Wrap a result as JSON text content."""
    return [TextContent(type="text", text=json.dumps(
        result, indent=2, ensure_ascii=False))]


def _error(msg: str) -> list[TextContent]:
    return [TextContent(type="text", text=json.dumps(
        {"error": msg}, indent=2, ensure_ascii=False))]


# ---------------------------------------------------------------------------
# Main dispatcher
# ---------------------------------------------------------------------------

async def handle_tool(name: str, arguments: dict | None) -> CallToolResult:
    """Dispatch tool calls to the registered handler or doc tool."""
    if arguments is None:
        arguments = {}

    log.info(f"Tool call: {name}")

    # ---- Documentation tools (no COM needed) ----
    if name == "list_docs":
        try:
            result = doc_tools.handle_list_docs()
            return CallToolResult(content=_json(result))
        except Exception as e:
            log.exception("list_docs failed")
            return CallToolResult(content=_error(str(e)))

    if name == "read_doc":
        try:
            result = doc_tools.handle_read_doc(arguments.get("name", ""))
            return CallToolResult(content=_json(result))
        except FileNotFoundError as e:
            return CallToolResult(content=_error(str(e)))
        except Exception as e:
            log.exception("read_doc failed")
            return CallToolResult(content=_error(str(e)))

    if name == "search_docs":
        try:
            result = doc_tools.handle_search_docs(arguments.get("query", ""))
            return CallToolResult(content=_json(result))
        except Exception as e:
            log.exception("search_docs failed")
            return CallToolResult(content=_error(str(e)))

    if name == "chm_to_html":
        try:
            result = await doc_tools.handle_convert_chm_to_html(
                arguments["chm_path"],
                arguments.get("output_dir"),
            )
            return CallToolResult(content=_json(result))
        except Exception as e:
            log.exception("chm_to_html failed")
            return CallToolResult(content=_error(str(e)))

    # ---- AlphaCAM COM tools ----
    try:
        acam = get_acam()
    except AlphaCAMNotRunning:
        return CallToolResult(content=_error(
            "AlphaCAM is not running. Start AlphaCAM first."))
    except Exception as e:
        return CallToolResult(content=_error(str(e)))

    func = get_tool_func(name)
    if func is None:
        return CallToolResult(content=_error(f"Unknown tool: {name}"))

    try:
        result = await func(acam, **arguments)
        if result is None:
            result = {"status": "ok"}
        return CallToolResult(content=_json(result))
    except ToolError as e:
        log.warning(f"Tool {name} error: {e.code} - {e}")
        return CallToolResult(content=_json(e.to_dict()))
    except AlphaCAMError as e:
        log.warning(f"AlphaCAM error in {name}: {e}")
        return CallToolResult(content=_error(str(e)))
    except KeyError as e:
        return CallToolResult(content=_error(
            f"Missing required argument: {e}"))
    except Exception as e:
        log.exception(f"Unexpected error in tool {name}")
        return CallToolResult(content=_error(str(e)))
