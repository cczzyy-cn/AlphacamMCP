"""
AlphaCAM MCP Server — exposes AlphaCAM 2016 R1 automation as MCP tools.

Each tool maps to a COM operation via the AlphaCAM wrapper.
The AI assistant can call these tools to read/modify/control AlphaCAM.

Usage:
    pip install mcp pywin32
    python server.py                 # stdio transport (for Reasonix)
    python server.py --port 8080     # SSE transport (for remote clients)
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from typing import Any

# ---------------------------------------------------------------------------
# MCP imports
# ---------------------------------------------------------------------------
try:
    from mcp.server import Server, NotificationOptions
    from mcp.server.models import InitializationOptions
    from mcp.types import (
        Tool,
        TextContent,
        ImageContent,
        EmbeddedResource,
        CallToolResult,
    )
    import mcp.server.stdio
except ImportError:
    sys.exit("mcp package not installed. Run: pip install mcp")

# ---------------------------------------------------------------------------
# AlphaCAM COM wrapper
# ---------------------------------------------------------------------------
# Add the skill directory to path so we can import alphacam_com
_skill_dir = os.path.dirname(os.path.abspath(__file__))
if _skill_dir not in sys.path:
    sys.path.insert(0, _skill_dir)

try:
    from alphacam_com import AlphaCAM, AlphaCAMError, AlphaCAMNotRunning
except ImportError:
    sys.exit(
        "alphacam_com.py not found. Make sure it's in the same directory."
    )

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger("alphacam-bridge")

# ---------------------------------------------------------------------------
# Global AlphaCAM instance
# ---------------------------------------------------------------------------
_acam: AlphaCAM | None = None
_prog_id: str = "aroutaps.Application"


def get_acam() -> AlphaCAM:
    """Get or create the shared AlphaCAM connection."""
    global _acam, _prog_id
    if _acam is None:
        _acam = AlphaCAM(prog_id=_prog_id)
    return _acam


# ---------------------------------------------------------------------------
# Tool definitions
# ---------------------------------------------------------------------------

TOOLS: list[dict] = [
    # ----- Status & Info -----
    {
        "name": "get_status",
        "description": "Check if AlphaCAM is running and return version/info.",
        "inputSchema": {
            "type": "object",
            "properties": {},
        },
    },
    {
        "name": "get_drawing_info",
        "description": "Get info about the active drawing: name, geo count, toolpath count, layers, operations.",
        "inputSchema": {
            "type": "object",
            "properties": {},
        },
    },
    # ----- File Operations -----
    {
        "name": "new_drawing",
        "description": "Create a new empty drawing (clears current).",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "open_drawing",
        "description": "Open an AlphaCAM drawing file (.amd).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Full path to the .amd file",
                }
            },
            "required": ["file_path"],
        },
    },
    {
        "name": "open_dxf",
        "description": "Open a DXF/DWG file.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "clear": {
                    "type": "boolean",
                    "description": "Clear current drawing first",
                    "default": True,
                },
            },
            "required": ["file_path"],
        },
    },
    {
        "name": "open_step",
        "description": "Open a STEP file.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "clear": {"type": "boolean", "default": True},
            },
            "required": ["file_path"],
        },
    },
    {
        "name": "open_stl",
        "description": "Open an STL file.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string"},
                "clear": {"type": "boolean", "default": True},
            },
            "required": ["file_path"],
        },
    },
    {
        "name": "save_drawing",
        "description": "Save the active drawing. Provide file_path to Save As.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Optional: Save As path",
                }
            },
        },
    },
    # ----- Geometry Creation -----
    {
        "name": "create_rectangle",
        "description": "Create a rectangle geometry.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "x1": {"type": "number", "description": "Bottom-left X"},
                "y1": {"type": "number", "description": "Bottom-left Y"},
                "x2": {"type": "number", "description": "Top-right X"},
                "y2": {"type": "number", "description": "Top-right Y"},
            },
            "required": ["x1", "y1", "x2", "y2"],
        },
    },
    {
        "name": "create_circle",
        "description": "Create a circle by diameter and center.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "diameter": {"type": "number"},
                "xc": {"type": "number", "description": "Center X"},
                "yc": {"type": "number", "description": "Center Y"},
            },
            "required": ["diameter", "xc", "yc"],
        },
    },
    {
        "name": "create_line",
        "description": "Create a 2D line between two points.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "x1": {"type": "number"}, "y1": {"type": "number"},
                "x2": {"type": "number"}, "y2": {"type": "number"},
            },
            "required": ["x1", "y1", "x2", "y2"],
        },
    },
    {
        "name": "create_polygon",
        "description": "Create a regular polygon.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "xc": {"type": "number", "description": "Center X"},
                "yc": {"type": "number", "description": "Center Y"},
                "radius": {"type": "number"},
                "sides": {"type": "integer", "description": "Number of sides"},
                "start_angle": {
                    "type": "number",
                    "description": "Start angle in degrees",
                    "default": 0,
                },
            },
            "required": ["xc", "yc", "radius", "sides"],
        },
    },
    {
        "name": "create_ellipse",
        "description": "Create an ellipse.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "xc": {"type": "number"}, "yc": {"type": "number"},
                "maj_ax": {"type": "number", "description": "Major axis length"},
                "min_ax": {"type": "number", "description": "Minor axis length"},
                "angle": {"type": "number", "description": "Rotation angle", "default": 0},
            },
            "required": ["xc", "yc", "maj_ax", "min_ax"],
        },
    },
    {
        "name": "create_text",
        "description": "Create a text annotation on the drawing.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "text": {"type": "string"},
                "height": {"type": "number"},
                "x": {"type": "number"}, "y": {"type": "number"},
                "font": {"type": "string", "default": "Arial"},
                "angle": {"type": "number", "default": 0},
            },
            "required": ["text", "height", "x", "y"],
        },
    },
    # ----- Work Plane & Layer -----
    {
        "name": "create_workplane",
        "description": "Create a named work plane with origin and orientation.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {"type": "string"},
                "x": {"type": "number", "default": 0},
                "y": {"type": "number", "default": 0},
                "z": {"type": "number", "default": 0},
                "i": {"type": "number", "default": 0, "description": "Z axis X component"},
                "j": {"type": "number", "default": 0, "description": "Z axis Y component"},
                "k": {"type": "number", "default": 1, "description": "Z axis Z component"},
            },
            "required": ["name"],
        },
    },
    {
        "name": "set_workplane",
        "description": "Set the active work plane by name.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {"type": "string"}
            },
            "required": ["name"],
        },
    },
    {
        "name": "create_layer",
        "description": "Create or get a layer by name, optionally set color.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {"type": "string"},
                "color": {
                    "type": "integer",
                    "description": "Optional RGB color value",
                },
            },
            "required": ["name"],
        },
    },
    # ----- Tool -----
    {
        "name": "select_tool",
        "description": "Select a tool from the library by path, or use '$USER' for dialog.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "name_or_path": {
                    "type": "string",
                    "description": "Tool file path or '$USER' for dialog",
                    "default": "$USER",
                }
            },
        },
    },
    {
        "name": "get_current_tool",
        "description": "Get info about the currently selected tool.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    # ----- Machining -----
    {
        "name": "run_machining",
        "description": "Run a machining operation with the given parameters.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "process_type": {
                    "type": "integer",
                    "description": "1=RoughFinish, 2=ContourPocket, 10=Engrave, 21=Drill, 22=Peck, 23=Tap, 24=Bore",
                    "default": 1,
                },
                "safe_rapid_level": {"type": "number", "default": 20},
                "rapid_down_to": {"type": "number", "default": 1},
                "final_depth": {"type": "number", "default": -5},
                "material_top": {"type": "number", "default": 0},
                "depth_of_cut": {"type": "number", "default": 2},
                "cut_feed": {"type": "number", "default": 1000},
                "down_feed": {"type": "number", "default": 500},
                "spindle_speed": {"type": "number", "default": 6000},
                "stock": {"type": "number", "default": 0},
                "width_of_cut": {"type": "number", "description": "Stepover for pocketing"},
                "bidirectional": {"type": "boolean", "default": False},
                "coolant": {"type": "integer", "default": 0, "description": "0=None, 1=Mist, 2=Flood, 3=Tool"},
                "mc_comp": {"type": "integer", "default": 0, "description": "0=ToolCenter, 1=MC, 2=Both"},
            },
        },
    },
    # ----- NC Output -----
    {
        "name": "output_nc",
        "description": "Output NC code to a file.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Full path for NC output file",
                },
                "visible_only": {
                    "type": "boolean",
                    "description": "Only output visible operations",
                    "default": True,
                },
            },
            "required": ["file_path"],
        },
    },
    # ----- VBA / Add-ins -----
    {
        "name": "run_vba_macro",
        "description": "Run a VBA macro by name with up to 8 optional parameters.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "macro_name": {"type": "string"},
                "params": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Up to 8 string parameters",
                    "maxItems": 8,
                },
            },
            "required": ["macro_name"],
        },
    },
    {
        "name": "load_addin",
        "description": "Load an add-in DLL or VBA project file.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_name": {
                    "type": "string",
                    "description": "Full path to the add-in file (.dll, .amb, etc.)",
                }
            },
            "required": ["file_name"],
        },
    },
    {
        "name": "enable_addin",
        "description": "Enable or disable an add-in by name.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {"type": "string"},
                "enable": {"type": "boolean", "default": True},
            },
            "required": ["name"],
        },
    },
    # ----- Utilities -----
    {
        "name": "set_undo_point",
        "description": "Set an undo point with a descriptive name.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "text": {"type": "string", "default": "AI Operation"}
            },
        },
    },
    {
        "name": "zoom_all",
        "description": "Zoom to extents in the active drawing.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "select_post",
        "description": "Select a post-processor by name.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "post_name": {"type": "string"}
            },
            "required": ["post_name"],
        },
    },
    {
        "name": "list_geometries",
        "description": "List all geometries in the active drawing.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "list_operations",
        "description": "List all operations in the active drawing.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "list_toolpaths",
        "description": "List all toolpaths in the active drawing.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "delete_selected",
        "description": "Delete all selected geometries in the active drawing.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "delete_all_geometries",
        "description": "Delete ALL geometries in the active drawing (keeps toolpaths, layers, etc.). Use with caution!",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "trim_with_boundary",
        "description": "Trim lines by a boundary (closed path). Breaks lines at boundary edges and removes segments outside the boundary. If boundary_index=0, uses the currently selected geometry as boundary. If line_indices is omitted, trims ALL lines that intersect the boundary.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "boundary_index": {
                    "type": "integer",
                    "description": "1-based index of boundary geometry (0 = use selected geometry)",
                    "default": 0,
                },
                "line_indices": {
                    "type": "array",
                    "items": {"type": "integer"},
                    "description": "1-based indices of lines to trim (omit to trim all intersecting lines)",
                },
            },
        },
    },
    {
        "name": "run_workflow",
        "description": "Run a batch of steps sequentially. Each step: {'action': 'method_name', 'params': {...}}. Use set_undo_point is called automatically.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "steps": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "action": {"type": "string"},
                            "params": {"type": "object"},
                        },
                        "required": ["action"],
                    },
                }
            },
            "required": ["steps"],
        },
    },

    {
        "name": "list_addins",
        "description": "List all loaded AlphaCAM add-ins with their connection status.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    # ----- Screen Locking -----
    {
        "name": "lock_acam",
        "description": "Disable screen redraw in AlphaCAM (for batch operations).",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "unlock_acam",
        "description": "Re-enable screen redraw and optionally zoom extents.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "zoom_all": {
                    "type": "boolean",
                    "description": "Zoom extents after unlock",
                    "default": False,
                }
            },
        },
    },
    # ----- Nesting -----
    {
        "name": "has_nesting",
        "description": "Check if the active drawing has nesting (NestInformation with sheets).",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "get_nesting_info",
        "description": "Get detailed nesting information: sheets, parts, instance data.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "get_sheet_extents",
        "description": "Get the global min/max X/Y extents across all nesting sheets.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    # ----- Path Operations -----
    {
        "name": "mirror_path",
        "description": "Mirror a path (geometry or toolpath) about a line defined by two points (MirrorL).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path_index": {"type": "integer", "description": "1-based index (0 = use selected)", "default": 0},
                "x1": {"type": "number", "description": "First point X on mirror line"},
                "y1": {"type": "number", "description": "First point Y on mirror line"},
                "x2": {"type": "number", "description": "Second point X on mirror line"},
                "y2": {"type": "number", "description": "Second point Y on mirror line"},
            },
            "required": ["x1", "y1", "x2", "y2"],
        },
    },
    {
        "name": "copy_temporary_store",
        "description": "Copy a path as temporary, optionally mirror it, then store it back (CopyTemporary + MirrorL + StoreTemporary).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path_index": {"type": "integer", "description": "1-based index (0 = use selected)", "default": 0},
                "mirror": {
                    "type": "object",
                    "description": "Optional mirror: {x1,y1,x2,y2}",
                    "properties": {
                        "x1": {"type": "number"},
                        "y1": {"type": "number"},
                        "x2": {"type": "number"},
                        "y2": {"type": "number"},
                    },
                },
            },
        },
    },
    {
        "name": "offset_path",
        "description": "Offset a closed path by a distance. side: 1=Left(outside), -1=Right(inside).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path_index": {"type": "integer", "description": "1-based index (0 = use selected)", "default": 0},
                "distance": {"type": "number", "description": "Offset distance"},
                "side": {"type": "integer", "description": "1=Left(outside), -1=Right(inside)", "default": 1},
                "delete_original": {"type": "boolean", "description": "Delete original path", "default": False},
            },
            "required": ["distance"],
        },
    },
    # ----- Attributes -----
    {
        "name": "get_path_attributes",
        "description": "Read all user attributes from a path (geometry or toolpath).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path_index": {"type": "integer", "description": "1-based index (0 = use selected)", "default": 0},
                "name_filter": {"type": "string", "description": "Optional substring filter for attribute names"},
            },
        },
    },
    {
        "name": "set_path_attribute",
        "description": "Set a user attribute on a path (geometry or toolpath).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path_index": {"type": "integer", "description": "1-based index (0 = use selected)", "default": 0},
                "attribute_name": {"type": "string", "description": "Attribute name"},
                "attribute_value": {"type": ["string", "number"], "description": "Attribute value"},
            },
            "required": ["attribute_name", "attribute_value"],
        },
    },
    # ----- Extended Queries -----
    {
        "name": "get_all_geometries",
        "description": "List ALL geometries in the active drawing with full details (type, extents, sheet/dim status, attributes).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "include_attributes": {
                    "type": "boolean",
                    "description": "Include user attributes for each geometry",
                    "default": False,
                },
            },
        },
    },

    # ----- Documentation -----
    {
        "name": "list_docs",
        "description": "List AlphaCAM API documentation categories and their file counts. Returns an overview of all available doc sections (General, Enums, Events, Objects, Post, Examples) with document counts.",
        "inputSchema": {
            "type": "object",
            "properties": {},
        },
    },
    {
        "name": "read_doc",
        "description": "Read an AlphaCAM API documentation page by name. Provide a filename (e.g. 'Path_TrimWithCuttingGeos', 'Drawing_CreateRectangle', 'Application_OpenDrawing') or a partial path. The full HTML content is returned as plain text. Browse docs with list_docs first.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Doc filename or identifier, e.g. 'Path_TrimWithCuttingGeos', 'DrawingObject', 'InitAlphacamAddIn'",
                }
            },
            "required": ["name"],
        },
    },
    {
        "name": "search_docs",
        "description": "Search AlphaCAM API documentation pages by keyword. Finds all pages whose filename or content matches the query. Returns matching file names with one-line summaries. Use read_doc to read a specific page.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search keyword, e.g. 'trim', 'offset', 'circle', 'drawing', 'tool'",
                }
            },
            "required": ["query"],
        },
    },
]


# ---------------------------------------------------------------------------
# Tool implementations
# ---------------------------------------------------------------------------

def _json(result: Any) -> list[TextContent]:
    """Wrap a result as JSON text content."""
    return [TextContent(type="text", text=json.dumps(result, indent=2,
                                                     ensure_ascii=False))]


def _error(msg: str) -> list[TextContent]:
    return [TextContent(type="text", text=json.dumps(
        {"error": msg}, indent=2, ensure_ascii=False))]


def _ensure_kw(**kwargs) -> dict:
    """Filter out None values from kwargs."""
    return {k: v for k, v in kwargs.items() if v is not None}


# ---------------------------------------------------------------------------
# Documentation helpers
# ---------------------------------------------------------------------------

_ALPHACAM_DIR = r"C:\Program Files (x86)\Vero Software\Alphacam 2016 R1"
_API_DOC_DIR = os.path.join(_ALPHACAM_DIR, "tempacamapi")

# Common CHM docs with descriptions
_CHM_DOCS = {
    "ACAMAPI": {"file": "ACAMAPI.chm", "desc": "API reference (VBA object model, methods, properties, events)"},
    "ACAM3": {"file": "ACAM3.chm", "desc": "3D module user manual"},
    "ACAM4": {"file": "ACAM4.chm", "desc": "4-axis module user manual"},
    "AEDIT3": {"file": "AEDIT3.chm", "desc": "3D editor help"},
    "AEDITAPI": {"file": "AEDITAPI.chm", "desc": "Editor API reference"},
    "AcamReports": {"file": "AcamReports.chm", "desc": "Reports system help"},
    "C2C": {"file": "C2C.chm", "desc": "CAD to CAM conversion help"},
    "ConstraintsAPI": {"file": "ConstraintsAPI.chm", "desc": "Constraints API reference"},
    "Feature": {"file": "Feature.chm", "desc": "Feature help"},
    "ModuleWorks": {"file": "ModuleWorks_-_Documentation.chm", "desc": "ModuleWorks machining engine docs (5-axis)"},
    "Primitives": {"file": "Primitives.chm", "desc": "Primitives help"},
    "R2V": {"file": "R2V.chm", "desc": "Raster to Vector help"},
    "simulate": {"file": "simulate.chm", "desc": "Simulator help"},
}


def _get_doc_categories() -> dict:
    """Return doc categories with file counts (recursive for Objects)."""
    categories = {}
    base = _API_DOC_DIR
    for entry in ["General", "Enums", "Events", "Objects", "Examples", "Post"]:
        path = os.path.join(base, entry)
        if os.path.isdir(path):
            count = 0
            for _root, _dirs, files in os.walk(path):
                count += sum(1 for f in files
                              if f.endswith(".htm") or f.endswith(".html"))
            categories[entry] = count
    return categories


def _find_doc_file(name: str) -> str | None:
    """Find a doc HTML file by name (case-insensitive, partial match)."""
    name_lower = name.lower()
    if not name_lower.endswith(".htm"):
        name_lower += ".htm"

    for root, _dirs, files in os.walk(_API_DOC_DIR):
        for f in files:
            if f.lower() == name_lower:
                return os.path.join(root, f)

    # Second pass: partial match
    for root, _dirs, files in os.walk(_API_DOC_DIR):
        for f in files:
            if f.lower().endswith(".htm") and name_lower in f.lower():
                return os.path.join(root, f)

    return None


def _strip_html(html: str) -> str:
    """Crude HTML-to-text conversion: strip tags, decode entities."""
    import re
    text = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'</?(?:p|div|br|tr|li|h\d|table|section)[^>]*>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = text.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
    text = text.replace('&nbsp;', ' ').replace('&#160;', ' ')
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'&#(\d+);', lambda m: chr(int(m.group(1))), text)
    return text.strip()


def _get_doc_title(filepath: str) -> str:
    """Extract the <title> from an HTML file."""
    import re
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            content = f.read(4096)
        m = re.search(r'<title[^>]*>(.*?)</title>', content, re.IGNORECASE | re.DOTALL)
        if m:
            return _strip_html(m.group(1))
    except Exception:
        pass
    return os.path.basename(filepath).replace(".htm", "").replace("_", " ")


def _search_docs(query: str, max_results: int = 20) -> list[dict]:
    """Search doc page filenames and titles for a query string."""
    q = query.lower()
    results = []
    for root, _dirs, files in os.walk(_API_DOC_DIR):
        for f in files:
            if not f.endswith(".htm"):
                continue
            filepath = os.path.join(root, f)
            rel_path = os.path.relpath(filepath, _API_DOC_DIR)
            score = 0
            if q in f.lower():
                score += 2
            title = _get_doc_title(filepath)
            if q in title.lower():
                score += 1
            if score > 0:
                results.append({
                    "file": f,
                    "title": title,
                    "path": rel_path,
                    "score": score,
                })
    results.sort(key=lambda x: -x["score"])
    return results[:max_results]


async def handle_list_docs() -> dict:
    """Handle the list_docs tool."""
    categories = _get_doc_categories()
    return {
        "api_docs": {
            "location": _API_DOC_DIR,
            "total_files": sum(categories.values()),
            "categories": categories,
        },
        "chm_docs": {k: v["desc"] for k, v in _CHM_DOCS.items()},
        "tip": "Use read_doc(name='Path_TrimWithCuttingGeos') to read a specific API page. Use search_docs(query='offset') to find docs by keyword.",
    }


async def handle_read_doc(name: str) -> dict:
    """Handle the read_doc tool."""
    filepath = _find_doc_file(name)
    if not filepath:
        raise FileNotFoundError(
            f"Document '{name}' not found. Use search_docs() to find matching pages."
        )
    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        html = f.read()
    text = _strip_html(html)
    rel_path = os.path.relpath(filepath, _API_DOC_DIR)
    title = _get_doc_title(filepath)
    MAX_LEN = 8000
    if len(text) > MAX_LEN:
        text = text[:MAX_LEN] + f"\n\n... [truncated, full length: {len(text)} chars]"
    return {
        "title": title,
        "file": os.path.basename(filepath),
        "path": rel_path,
        "content": text,
    }


async def handle_search_docs(query: str) -> dict:
    """Handle the search_docs tool."""
    results = _search_docs(query)
    if not results:
        return {
            "query": query,
            "count": 0,
            "results": [],
            "tip": "Try a different keyword, or use list_docs() to browse categories.",
        }
    return {
        "query": query,
        "count": len(results),
        "results": [
            {"file": r["file"], "title": r["title"], "path": r["path"]}
            for r in results
        ],
    }


async def handle_tool(name: str, arguments: dict | None) -> CallToolResult:
    """Dispatch tool calls to the AlphaCAM wrapper."""
    if arguments is None:
        arguments = {}

    acam = get_acam()
    result: Any = None

    try:
        log.info(f"Tool call: {name} {arguments}")

        # Status & Info
        if name == "get_status":
            result = acam.get_info()
        elif name == "get_drawing_info":
            result = acam.get_drawing_info()

        # File
        elif name == "new_drawing":
            acam.new_drawing()
            result = {"status": "ok", "message": "New drawing created"}
        elif name == "open_drawing":
            result = acam.open_drawing(arguments["file_path"])
        elif name == "open_dxf":
            result = acam.open_dxf(arguments["file_path"],
                                   arguments.get("clear", True))
        elif name == "open_step":
            result = acam.open_step(arguments["file_path"],
                                    arguments.get("clear", True))
        elif name == "open_stl":
            result = acam.open_stl(arguments["file_path"],
                                   arguments.get("clear", True))
        elif name == "save_drawing":
            acam.save_drawing(arguments.get("file_path"))
            result = {"status": "ok"}

        # Geometry
        elif name == "create_rectangle":
            result = acam.create_rectangle(
                arguments["x1"], arguments["y1"],
                arguments["x2"], arguments["y2"])
        elif name == "create_circle":
            result = acam.create_circle(
                arguments["diameter"],
                arguments["xc"], arguments["yc"])
        elif name == "create_line":
            result = acam.create_line(
                arguments["x1"], arguments["y1"],
                arguments["x2"], arguments["y2"])
        elif name == "create_polygon":
            result = acam.create_polygon(
                arguments["xc"], arguments["yc"],
                arguments["radius"], arguments["sides"],
                arguments.get("start_angle", 0))
        elif name == "create_ellipse":
            result = acam.create_ellipse(
                arguments["xc"], arguments["yc"],
                arguments["maj_ax"], arguments["min_ax"],
                arguments.get("angle", 0))
        elif name == "create_text":
            result = acam.create_text(
                arguments["text"], arguments["height"],
                arguments["x"], arguments["y"],
                arguments.get("font", "Arial"),
                arguments.get("angle", 0))

        # Workplane & Layer
        elif name == "create_workplane":
            result = acam.create_workplane(
                arguments["name"],
                arguments.get("x", 0), arguments.get("y", 0),
                arguments.get("z", 0),
                arguments.get("i", 0), arguments.get("j", 0),
                arguments.get("k", 1))
        elif name == "set_workplane":
            acam.set_workplane(arguments["name"])
            result = {"status": "ok", "workplane": arguments["name"]}
        elif name == "create_layer":
            result = acam.create_layer(
                arguments["name"], arguments.get("color"))

        # Tool
        elif name == "select_tool":
            result = acam.select_tool(arguments.get("name_or_path", "$USER"))
        elif name == "get_current_tool":
            result = acam.get_current_tool()

        # Machining
        elif name == "run_machining":
            result = acam.run_machining(_ensure_kw(**arguments))

        # NC
        elif name == "output_nc":
            result = acam.output_nc(
                arguments["file_path"],
                visible_only=arguments.get("visible_only", True))

        # VBA / Add-ins
        elif name == "run_vba_macro":
            result = acam.run_vba_macro(
                arguments["macro_name"],
                arguments.get("params"))
        elif name == "load_addin":
            result = acam.load_addin(arguments["file_name"])
        elif name == "enable_addin":
            result = acam.enable_addin(
                arguments["name"], arguments.get("enable", True))

        # Utilities
        elif name == "set_undo_point":
            result = acam.set_undo_point(arguments.get("text", "AI Operation"))
        elif name == "zoom_all":
            acam.zoom_all()
            result = {"status": "ok"}
        elif name == "select_post":
            result = acam.select_post(arguments["post_name"])
        elif name == "list_geometries":
            result = acam.list_geometries()
        elif name == "list_operations":
            result = acam.list_operations()
        elif name == "list_toolpaths":
            result = acam.list_toolpaths()

        # Delete
        elif name == "delete_selected":
            result = acam.delete_selected()
        elif name == "delete_all_geometries":
            result = acam.delete_all_geometries()
        elif name == "trim_with_boundary":
            result = acam.trim_with_boundary(
                boundary_index=arguments.get("boundary_index", 0),
                line_indices=arguments.get("line_indices"),
            )

        # Batch
        elif name == "run_workflow":
            result = acam.run_workflow(arguments["steps"])


        elif name == "list_addins":
            result = acam.list_addins()
        # Screen Locking
        elif name == "lock_acam":
            result = acam.lock_acam()
        elif name == "unlock_acam":
            result = acam.unlock_acam(arguments.get("zoom_all", False))
        # Nesting
        elif name == "has_nesting":
            result = acam.has_nesting()
        elif name == "get_nesting_info":
            result = acam.get_nesting_info()
        elif name == "get_sheet_extents":
            result = acam.get_sheet_extents()
        # Path Operations
        elif name == "mirror_path":
            result = acam.mirror_path(
                arguments["x1"], arguments["y1"],
                arguments["x2"], arguments["y2"],
                arguments.get("path_index", 0))
        elif name == "copy_temporary_store":
            result = acam.copy_temporary_store(
                arguments.get("path_index", 0),
                arguments.get("mirror"))
        elif name == "offset_path":
            result = acam.offset_path(
                arguments["distance"],
                arguments.get("side", 1),
                arguments.get("path_index", 0),
                arguments.get("delete_original", False))
        # Attributes
        elif name == "get_path_attributes":
            result = acam.get_path_attributes(
                arguments.get("path_index", 0),
                arguments.get("name_filter"))
        elif name == "set_path_attribute":
            result = acam.set_path_attribute(
                arguments["attribute_name"],
                arguments["attribute_value"],
                arguments.get("path_index", 0))
        # Extended Queries
        elif name == "get_all_geometries":
            result = acam.get_all_geometries(
                arguments.get("include_attributes", False))

        # Documentation
        elif name == "list_docs":
            result = await handle_list_docs()
        elif name == "read_doc":
            result = await handle_read_doc(arguments["name"])
        elif name == "search_docs":
            result = await handle_search_docs(arguments["query"])

        else:
            return CallToolResult(
                content=_error(f"Unknown tool: {name}"),
                isError=True,
            )

        return CallToolResult(content=_json(result))

    except AlphaCAMNotRunning as exc:
        return CallToolResult(
            content=_error(f"AlphaCAM not available: {exc}"),
            isError=True,
        )
    except FileNotFoundError as exc:
        return CallToolResult(
            content=_error(f"File not found: {exc}"),
            isError=True,
        )
    except AlphaCAMError as exc:
        return CallToolResult(
            content=_error(f"AlphaCAM error: {exc}"),
            isError=True,
        )
    except Exception as exc:
        log.exception(f"Unhandled error in {name}")
        return CallToolResult(
            content=_error(f"Unexpected error: {exc}"),
            isError=True,
        )


# ---------------------------------------------------------------------------
# Server setup
# ---------------------------------------------------------------------------

def create_server() -> Server:
    """Create and configure the MCP server."""
    server = Server("alphacam-bridge")

    @server.list_tools()
    async def list_tools() -> list[Tool]:
        return [Tool(**t) for t in TOOLS]

    @server.call_tool()
    async def call_tool(name: str, arguments: dict | None) -> CallToolResult:
        return await handle_tool(name, arguments)

    return server


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

async def run_stdio():
    """Run via stdio transport (for Reasonix local integration)."""
    server = create_server()
    async with mcp.server.stdio.stdio_server() as (read, write):
        await server.run(
            read, write,
            InitializationOptions(
                server_name="alphacam-bridge",
                server_version="1.0.0",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )


def main():
    parser = argparse.ArgumentParser(
        description="AlphaCAM MCP Bridge Server")
    parser.add_argument("--port", type=int, default=None,
                        help="Port for SSE transport (omit for stdio)")
    parser.add_argument("--progid", type=str, default="aroutaps.Application",
                        help="COM ProgID (aroutaps.Application, am5axaps.Application, etc.)")
    args = parser.parse_args()

    global _prog_id
    _prog_id = args.progid

    if args.port:
        # SSE mode — for remote or web-based MCP clients
        from mcp.server.sse import SseServerTransport
        from starlette.applications import Starlette
        from starlette.routing import Mount, Route
        import uvicorn

        server = create_server()
        sse = SseServerTransport("/messages/")

        async def handle_sse(request):
            async with sse.connect_sse(
                request.scope, request.receive, request._send,
            ) as (read, write):
                await server.run(
                    read, write,
                    InitializationOptions(
                        server_name="alphacam-bridge",
                        server_version="1.0.0",
                        capabilities=server.get_capabilities(
                            notification_options=NotificationOptions(),
                            experimental_capabilities={},
                        ),
                    ),
                )

        app = Starlette(
            routes=[
                Route("/sse", endpoint=handle_sse),
                Mount("/messages/", app=sse.handle_post_message),
            ],
        )
        log.info(f"Starting SSE server on port {args.port}")
        uvicorn.run(app, host="127.0.0.1", port=args.port)
    else:
        # stdio mode
        import asyncio
        log.info("Starting stdio server")
        asyncio.run(run_stdio())


if __name__ == "__main__":
    main()
