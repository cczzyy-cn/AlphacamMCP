"""
AlphaCAM MCP Bridge — modular package.

Provides MCP tools to control AlphaCAM 2016 R1 via COM automation.
"""

from .tools import TOOLS
from .handler import handle_tool, get_acam, set_prog_id
from .config import get_prog_id, get_visible, detect_alphacam_dir
from .errors import ToolError

__all__ = [
    "TOOLS", "handle_tool", "get_acam", "set_prog_id",
    "get_prog_id", "get_visible", "detect_alphacam_dir",
    "ToolError",
]
