"""
MCP tool dispatcher — routes tool calls to registered handlers.
"""

from __future__ import annotations

import gc
import json
import logging
from typing import Any

from mcp.types import TextContent, CallToolResult

from alphacam_com import AlphaCAM, AlphaCAMError, AlphaCAMNotRunning, ConnectionState
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
        # Wire up a logging callback for connection state changes
        _acam.set_state_callback(_on_connection_state_change)
    return _acam


def _on_connection_state_change(
    new_state: "ConnectionState", old_state: "ConnectionState"
):
    """Log connection state transitions and take recovery actions."""
    if new_state.name == "CONNECTED" and old_state.name != "CONNECTING":
        log.info("AlphaCAM connection restored (was %s)", old_state.value)
    elif new_state.name == "RECONNECTING":
        log.warning("AlphaCAM connection lost — attempting reconnect …")
    elif new_state.name == "FAILED":
        log.error("AlphaCAM connection FAILED after all retries")


def _get_acam_instance() -> AlphaCAM | None:
    """Return current acam instance without creating one (for config detection)."""
    return _acam


def set_prog_id(prog_id: str):
    """Override the ProgID (called from CLI --progid).

    Resets the current connection; the next ``get_acam()`` call will
    create a fresh AlphaCAM instance with the new ProgID.
    """
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
            expand = arguments.get("expand", False)
            result = doc_tools.handle_list_docs(expand=expand)
            return CallToolResult(content=_json(result))
        except Exception as e:
            log.exception("list_docs failed")
            return CallToolResult(content=_error(str(e)))

    if name == "read_doc":
        try:
            result = doc_tools.handle_read_doc(
                arguments.get("name", ""),
                max_len=arguments.get("max_len", 8000),
            )
            return CallToolResult(content=_json(result))
        except FileNotFoundError as e:
            return CallToolResult(content=_error(str(e)))
        except Exception as e:
            log.exception("read_doc failed")
            return CallToolResult(content=_error(str(e)))

    if name == "search_docs":
        try:
            result = doc_tools.handle_search_docs(
                arguments.get("query", ""),
                search_content=arguments.get("search_content", False),
            )
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

    if name == "chm_to_html_all":
        try:
            result = await doc_tools.handle_convert_all_chm(
                arguments.get("output_base_dir"),
            )
            return CallToolResult(content=_json(result))
        except Exception as e:
            log.exception("chm_to_html_all failed")
            return CallToolResult(content=_error(str(e)))

    # ---- AlphaCAM COM tools ----
    try:
        acam = get_acam()
        # Verify the COM connection is alive; reconnect if AlphaCAM was restarted
        acam.ensure_connection()
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
        err_code = getattr(e, 'hresult', 0) or getattr(e, 'args', [0])[0]
        is_zombie = (
            isinstance(err_code, int) and err_code in (
                -2147417851,  # 0x8000FFFF E_UNEXPECTED / catastrophic
                -2147221021,  # 0x800401F3 操作不可用
                -2147023174,  # 0x800706BA RPC 服务器不可用
            )
        )
        if is_zombie:
            log.warning(
                "COM zombie detected (code %d) — restarting connection …",
                err_code,
            )
            try:
                acam.restart()
                # Retry once with fresh connection
                result = await func(acam, **arguments)
                if result is None:
                    result = {"status": "ok"}
                return CallToolResult(content=_json(result))
            except Exception as retry_err:
                log.exception(f"Reconnect+retry of {name} also failed")
                return CallToolResult(content=_error(
                    f"AlphaCAM connection lost and reconnect failed: {retry_err}"))
        log.exception(f"Unexpected error in tool {name}")
        return CallToolResult(content=_error(str(e)))
    finally:
        # 释放 COM 临时引用，避免文件被锁定无法关闭
        gc.collect()
