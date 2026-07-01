"""
MCP tool definitions — registry + decorator + all 52 tool handler functions.
"""

from __future__ import annotations

import asyncio
import inspect
import logging
from typing import Any, Callable

from .errors import ToolComError

log = logging.getLogger("alphacam-bridge.tools")

# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

_tool_registry: dict[str, Callable] = {}
TOOLS: list[dict] = []


def _get_tool_param_names(func: Callable) -> list[str]:
    """Return param names excluding 'self', 'acam', and **kwargs."""
    sig = inspect.signature(func)
    names = []
    for name, param in sig.parameters.items():
        if name in ("self", "acam"):
            continue
        if param.kind == inspect.Parameter.VAR_KEYWORD:
            continue
        names.append(name)
    return names


def _get_tool_required_params(func: Callable) -> list[str]:
    """Return param names that have no default value (required)."""
    sig = inspect.signature(func)
    required = []
    for name, param in sig.parameters.items():
        if name in ("self", "acam", "kw", "kwargs"):
            continue
        if param.kind == inspect.Parameter.VAR_KEYWORD:
            continue
        if param.default is inspect.Parameter.empty:
            required.append(name)
    return required


def mcp_tool(tool_name: str, description: str, **properties):
    """Decorator: register an MCP tool with its input schema.

    Usage:
        @mcp_tool("create_circle", "Create a circle.",
            diameter={"type": "number", "description": "Diameter"},
            xc={"type": "number", "description": "Center X"},
            yc={"type": "number", "description": "Center Y"},
        )
        async def handle_create_circle(acam, diameter, xc, yc, **kw):
            return acam.create_circle(diameter, xc, yc)
    """
    def decorator(func):
        inspect.signature(func)
        required = _get_tool_required_params(func)

        _tool_registry[tool_name] = func
        TOOLS.append({
            "name": tool_name,
            "description": description,
            "inputSchema": {
                "type": "object",
                "properties": properties,
                "required": required,
            },
        })
        return func
    return decorator


def get_tool_func(name: str) -> Callable | None:
    return _tool_registry.get(name)


# ---------------------------------------------------------------------------
# Documentation tools — registered here but handled separately in handler.py
# (they don't need an AlphaCAM connection)
# ---------------------------------------------------------------------------

_DOC_TOOLS = [
    {
        "name": "list_docs",
        "description": "List AlphaCAM API documentation categories and their file counts. Returns an overview of all available doc sections (General, Enums, Events, Objects, Post, Examples) with document counts.",
        "inputSchema": {"type": "object", "properties": {}},
    },
    {
        "name": "read_doc",
        "description": "Read an AlphaCAM API documentation page by name. Provide a filename (e.g. 'Path_TrimWithCuttingGeos', 'Drawing_CreateRectangle', 'Application_OpenDrawing') or a partial path. The full HTML content is returned as plain text. Browse docs with list_docs first.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "Doc filename or identifier, e.g. 'Path_TrimWithCuttingGeos', 'DrawingObject', 'InitAlphacamAddIn'"},
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
                "query": {"type": "string", "description": "Search keyword, e.g. 'trim', 'offset', 'circle', 'drawing', 'tool'"},
            },
            "required": ["query"],
        },
    },
    {
        "name": "chm_to_html",
        "description": "Convert a .chm (Compiled HTML Help) file to HTML files using hh.exe. Extracts all pages, images, and assets to an output directory.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "chm_path": {"type": "string", "description": "Full path to the .chm file to convert"},
                "output_dir": {"type": "string", "description": "Optional output directory path. If omitted, creates a folder next to the .chm file with the same base name plus '_html'"},
            },
            "required": ["chm_path"],
        },
    },
]

TOOLS.extend(_DOC_TOOLS)


# ---------------------------------------------------------------------------
# COM retry helper
# ---------------------------------------------------------------------------

async def com_call(func, *args, max_retries: int = 2, delay: float = 0.5,
                   **kwargs) -> Any:
    """Call a COM method with automatic retry on transient errors."""
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if attempt < max_retries - 1:
                log.warning(f"COM call failed (attempt {attempt + 1}): {e}")
                await asyncio.sleep(delay)
            else:
                raise ToolComError(str(e), retries=attempt)


# ---------------------------------------------------------------------------
# Helper: consistent JSON result for "ok" responses
# ---------------------------------------------------------------------------

def _ok(**extra) -> dict:
    return {"status": "ok", **extra}


# ===========================================================================
# Tool handlers — each decorated with @mcp_tool
# All accept `acam` as first arg (may be None for doc-only tools).
# Return a dict (JSON-serializable) or None (becomes {"status": "ok"}).
# ===========================================================================

# ----- Status & Info (2) --------------------------------------------------

@mcp_tool("get_status",
    "Check if AlphaCAM is running and return version/info.")
async def handle_get_status(acam, **kw):
    return acam.get_info()

@mcp_tool("get_drawing_info",
    "Get info about the active drawing: name, geo count, toolpath count, layers, operations.")
async def handle_get_drawing_info(acam, **kw):
    return acam.get_drawing_info()

# ----- File Operations (7) ------------------------------------------------

@mcp_tool("new_drawing",
    "Create a new empty drawing (clears current).")
async def handle_new_drawing(acam, **kw):
    acam.new_drawing()
    return _ok(message="New drawing created")

@mcp_tool("open_drawing",
    "Open an AlphaCAM drawing file (.amd).",
    file_path={"type": "string", "description": "Full path to the .amd file"},
)
async def handle_open_drawing(acam, file_path, **kw):
    return acam.open_drawing(file_path)

@mcp_tool("open_dxf",
    "Open a DXF/DWG file.",
    file_path={"type": "string", "description": "Full path to the DXF file"},
    clear={"type": "boolean", "description": "Clear current drawing first"},
)
async def handle_open_dxf(acam, file_path, clear=True, **kw):
    return acam.open_dxf(file_path, clear)

@mcp_tool("open_step",
    "Open a STEP file.",
    file_path={"type": "string", "description": "Full path to the STEP file"},
    clear={"type": "boolean", "description": "Clear current drawing first"},
)
async def handle_open_step(acam, file_path, clear=True, **kw):
    return acam.open_step(file_path, clear)

@mcp_tool("open_stl",
    "Open an STL file.",
    file_path={"type": "string", "description": "Full path to the STL file"},
    clear={"type": "boolean", "description": "Clear current drawing first"},
)
async def handle_open_stl(acam, file_path, clear=True, **kw):
    return acam.open_stl(file_path, clear)

@mcp_tool("save_drawing",
    "Save the active drawing. Provide file_path to Save As.",
    file_path={"type": "string", "description": "Optional: Save As path"},
)
async def handle_save_drawing(acam, file_path=None, **kw):
    acam.save_drawing(file_path)
    return _ok()

@mcp_tool("output_nc",
    "Output NC code to a file.",
    file_path={"type": "string", "description": "Full path for NC output file"},
    visible_only={"type": "boolean", "description": "Only output visible operations"},
)
async def handle_output_nc(acam, file_path, visible_only=True, **kw):
    return acam.output_nc(file_path, visible_only)

# ----- Geometry Creation (6) ----------------------------------------------

@mcp_tool("create_rectangle",
    "Create a rectangle geometry.",
    x1={"type": "number", "description": "Bottom-left X"},
    y1={"type": "number", "description": "Bottom-left Y"},
    x2={"type": "number", "description": "Top-right X"},
    y2={"type": "number", "description": "Top-right Y"},
)
async def handle_create_rectangle(acam, x1, y1, x2, y2, **kw):
    return acam.create_rectangle(x1, y1, x2, y2)

@mcp_tool("create_circle",
    "Create a circle by diameter and center.",
    diameter={"type": "number", "description": "Circle diameter"},
    xc={"type": "number", "description": "Center X"},
    yc={"type": "number", "description": "Center Y"},
)
async def handle_create_circle(acam, diameter, xc, yc, **kw):
    return acam.create_circle(diameter, xc, yc)

@mcp_tool("create_line",
    "Create a 2D line between two points.",
    x1={"type": "number", "description": "Start point X"},
    y1={"type": "number", "description": "Start point Y"},
    x2={"type": "number", "description": "End point X"},
    y2={"type": "number", "description": "End point Y"},
)
async def handle_create_line(acam, x1, y1, x2, y2, **kw):
    return acam.create_line(x1, y1, x2, y2)

@mcp_tool("create_polygon",
    "Create a regular polygon.",
    radius={"type": "number", "description": "Circumscribed radius"},
    sides={"type": "integer", "description": "Number of sides"},
    xc={"type": "number", "description": "Center X"},
    yc={"type": "number", "description": "Center Y"},
    start_angle={"type": "number", "description": "Start angle in degrees"},
)
async def handle_create_polygon(acam, radius, sides, xc, yc, start_angle=0, **kw):
    return acam.create_polygon(radius, sides, xc, yc, start_angle)

@mcp_tool("create_ellipse",
    "Create an ellipse.",
    maj_ax={"type": "number", "description": "Major axis length"},
    min_ax={"type": "number", "description": "Minor axis length"},
    xc={"type": "number", "description": "Center X"},
    yc={"type": "number", "description": "Center Y"},
    angle={"type": "number", "description": "Rotation angle"},
)
async def handle_create_ellipse(acam, maj_ax, min_ax, xc, yc, angle=0, **kw):
    return acam.create_ellipse(maj_ax, min_ax, xc, yc, angle)

@mcp_tool("create_text",
    "Create a text annotation on the drawing.",
    text={"type": "string", "description": "Text content"},
    x={"type": "number", "description": "Position X"},
    y={"type": "number", "description": "Position Y"},
    height={"type": "number", "description": "Text height"},
    font={"type": "string", "description": "Font name"},
    angle={"type": "number", "description": "Rotation angle"},
)
async def handle_create_text(acam, text, x, y, height, font="Arial", angle=0, **kw):
    return acam.create_text(text, x, y, height, font, angle)

# ----- Work Plane & Layer (3) ---------------------------------------------

@mcp_tool("create_workplane",
    "Create a named work plane with origin and orientation.",
    name={"type": "string", "description": "Work plane name"},
    x={"type": "number", "description": "Origin X"},
    y={"type": "number", "description": "Origin Y"},
    z={"type": "number", "description": "Origin Z"},
    i={"type": "number", "description": "Z axis X component"},
    j={"type": "number", "description": "Z axis Y component"},
    k={"type": "number", "description": "Z axis Z component"},
)
async def handle_create_workplane(acam, name, x=0, y=0, z=0, i=0, j=0, k=1, **kw):
    return acam.create_workplane(name, x, y, z, i, j, k)

@mcp_tool("set_workplane",
    "Set the active work plane by name.",
    name={"type": "string", "description": "Work plane name"},
)
async def handle_set_workplane(acam, name, **kw):
    return acam.set_workplane(name)

@mcp_tool("create_layer",
    "Create or get a layer by name, optionally set color.",
    name={"type": "string", "description": "Layer name"},
    color={"type": "integer", "description": "Optional RGB color value"},
)
async def handle_create_layer(acam, name, color=None, **kw):
    return acam.create_layer(name, color)

# ----- Tool (2) -----------------------------------------------------------

@mcp_tool("select_tool",
    "Select a tool from the library by path, or use '$USER' for dialog.",
    name_or_path={"type": "string", "description": "Tool file path or '$USER' for dialog"},
)
async def handle_select_tool(acam, name_or_path="$USER", **kw):
    return acam.select_tool(name_or_path)

@mcp_tool("get_current_tool",
    "Get info about the currently selected tool.")
async def handle_get_current_tool(acam, **kw):
    return acam.get_current_tool()

# ----- Machining (1) ------------------------------------------------------

@mcp_tool("run_machining",
    "Run a machining operation with the given parameters.",
    process_type={"type": "integer", "description": "1=RoughFinish, 2=ContourPocket, 10=Engrave, 21=Drill, 22=Peck, 23=Tap, 24=Bore"},
    cut_feed={"type": "number", "description": "Cut feed rate"},
    down_feed={"type": "number", "description": "Plunge feed rate"},
    spindle_speed={"type": "number", "description": "Spindle speed (RPM)"},
    depth_of_cut={"type": "number", "description": "Depth per pass"},
    final_depth={"type": "number", "description": "Final depth (negative = below top)"},
    material_top={"type": "number", "description": "Material top Z"},
    safe_rapid_level={"type": "number", "description": "Safe rapid Z level"},
    rapid_down_to={"type": "number", "description": "Rapid down to Z"},
    stock={"type": "number", "description": "Stock allowance"},
    width_of_cut={"type": "number", "description": "Stepover for pocketing"},
    mc_comp={"type": "integer", "description": "0=ToolCenter, 1=MC, 2=Both"},
    bidirectional={"type": "boolean", "description": "Bidirectional cutting"},
    coolant={"type": "integer", "description": "0=None, 1=Mist, 2=Flood, 3=Tool"},
)
async def handle_run_machining(acam, process_type=1, cut_feed=1000, down_feed=500,
                                spindle_speed=6000, depth_of_cut=2, final_depth=-5,
                                material_top=0, safe_rapid_level=20, rapid_down_to=1,
                                stock=0, width_of_cut=None, mc_comp=0,
                                bidirectional=False, coolant=0, **kw):
    return acam.run_machining(
        process_type=process_type, cut_feed=cut_feed, down_feed=down_feed,
        spindle_speed=spindle_speed, depth_of_cut=depth_of_cut,
        final_depth=final_depth, material_top=material_top,
        safe_rapid_level=safe_rapid_level, rapid_down_to=rapid_down_to,
        stock=stock, width_of_cut=width_of_cut, mc_comp=mc_comp,
        bidirectional=bidirectional, coolant=coolant,
    )

# ----- VBA & AddIn (4) ----------------------------------------------------

@mcp_tool("run_vba_macro",
    "Run a VBA macro by name with up to 8 optional parameters.",
    macro_name={"type": "string", "description": "Macro name (Project.Module.Procedure)"},
    params={"type": "array", "items": {"type": "string"}, "description": "Up to 8 string parameters"},
)
async def handle_run_vba_macro(acam, macro_name, params=None, **kw):
    return acam.run_vba_macro(macro_name, params or [])

@mcp_tool("load_addin",
    "Load an add-in DLL or VBA project file.",
    file_name={"type": "string", "description": "Full path to the add-in file (.dll, .amb, etc.)"},
)
async def handle_load_addin(acam, file_name, **kw):
    return acam.load_addin(file_name)

@mcp_tool("enable_addin",
    "Enable or disable an add-in by name.",
    name={"type": "string", "description": "Add-in name"},
    enable={"type": "boolean", "description": "Enable or disable"},
)
async def handle_enable_addin(acam, name, enable=True, **kw):
    return acam.enable_addin(name, enable)

@mcp_tool("list_addins",
    "List all loaded AlphaCAM add-ins with their connection status.")
async def handle_list_addins(acam, **kw):
    return acam.list_addins()

# ----- Post Processor (1) -------------------------------------------------

@mcp_tool("select_post",
    "Select a post-processor by name.",
    post_name={"type": "string", "description": "Post-processor name"},
)
async def handle_select_post(acam, post_name, **kw):
    return acam.select_post(post_name)

# ----- Listing/Query (4) --------------------------------------------------

@mcp_tool("list_geometries",
    "List all geometries in the active drawing.")
async def handle_list_geometries(acam, **kw):
    return acam.list_geometries()

@mcp_tool("list_operations",
    "List all operations in the active drawing.")
async def handle_list_operations(acam, **kw):
    return acam.list_operations()

@mcp_tool("list_toolpaths",
    "List all toolpaths in the active drawing.")
async def handle_list_toolpaths(acam, **kw):
    return acam.list_toolpaths()

@mcp_tool("get_all_geometries",
    "List ALL geometries in the active drawing with full details (type, extents, sheet/dim status, attributes).",
    include_attributes={"type": "boolean", "description": "Include user attributes for each geometry"},
)
async def handle_get_all_geometries(acam, include_attributes=False, **kw):
    return acam.get_all_geometries(include_attributes)

# ----- Delete (2) ---------------------------------------------------------

@mcp_tool("delete_selected",
    "Delete all selected geometries in the active drawing.")
async def handle_delete_selected(acam, **kw):
    return acam.delete_selected()

@mcp_tool("delete_all_geometries",
    "Delete ALL geometries in the active drawing (keeps toolpaths, layers, etc.). Use with caution!")
async def handle_delete_all_geometries(acam, **kw):
    return acam.delete_all_geometries()

# ----- Trim (1) ------------------------------------------------------------

@mcp_tool("trim_with_boundary",
    "Trim lines by a boundary (closed path). Breaks lines at boundary edges and removes segments outside.",
    boundary_index={"type": "integer", "description": "1-based index of boundary geometry (0 = use selected)"},
    line_indices={"type": "array", "items": {"type": "integer"}, "description": "1-based indices of lines to trim (omit to trim all intersecting)"},
)
async def handle_trim_with_boundary(acam, boundary_index=0, line_indices=None, **kw):
    return acam.trim_with_boundary(boundary_index, line_indices)

# ----- Screen Control (2) --------------------------------------------------

@mcp_tool("lock_acam",
    "Disable screen redraw in AlphaCAM (for batch operations).")
async def handle_lock_acam(acam, **kw):
    return acam.lock_acam()

@mcp_tool("unlock_acam",
    "Re-enable screen redraw and optionally zoom extents.",
    zoom_all={"type": "boolean", "description": "Zoom extents after unlock"},
)
async def handle_unlock_acam(acam, zoom_all=False, **kw):
    return acam.unlock_acam(zoom_all)

# ----- Nesting (3) ---------------------------------------------------------

@mcp_tool("has_nesting",
    "Check if the active drawing has nesting (NestInformation with sheets).")
async def handle_has_nesting(acam, **kw):
    return acam.has_nesting()

@mcp_tool("get_nesting_info",
    "Get detailed nesting information: sheets, parts, instance data.")
async def handle_get_nesting_info(acam, **kw):
    return acam.get_nesting_info()

@mcp_tool("get_sheet_extents",
    "Get the global min/max X/Y extents across all nesting sheets.")
async def handle_get_sheet_extents(acam, **kw):
    return acam.get_sheet_extents()

# ----- Path Operations (5) -------------------------------------------------

@mcp_tool("mirror_path",
    "Mirror a path (geometry or toolpath) about a line defined by two points.",
    x1={"type": "number", "description": "First point X on mirror line"},
    y1={"type": "number", "description": "First point Y on mirror line"},
    x2={"type": "number", "description": "Second point X on mirror line"},
    y2={"type": "number", "description": "Second point Y on mirror line"},
    path_index={"type": "integer", "description": "1-based index (0 = use selected)"},
)
async def handle_mirror_path(acam, x1, y1, x2, y2, path_index=0, **kw):
    return acam.mirror_path(x1, y1, x2, y2, path_index)

@mcp_tool("offset_path",
    "Offset a closed path by a distance.",
    distance={"type": "number", "description": "Offset distance"},
    side={"type": "integer", "description": "1=Left(outside), -1=Right(inside)"},
    path_index={"type": "integer", "description": "1-based index (0 = use selected)"},
    delete_original={"type": "boolean", "description": "Delete original path"},
)
async def handle_offset_path(acam, distance, side=1, path_index=0, delete_original=False, **kw):
    return acam.offset_path(distance, side, path_index, delete_original)

@mcp_tool("copy_temporary_store",
    "Copy a path as temporary, optionally mirror it, then store it back.",
    path_index={"type": "integer", "description": "1-based index (0 = use selected)"},
    mirror={"type": "object", "description": "Optional mirror: {x1,y1,x2,y2}",
            "properties": {"x1": {"type": "number"}, "x2": {"type": "number"},
                        "y1": {"type": "number"}, "y2": {"type": "number"}}},
)
async def handle_copy_temporary_store(acam, path_index=0, mirror=None, **kw):
    return acam.copy_temporary_store(path_index, mirror)

@mcp_tool("get_path_attributes",
    "Read all user attributes from a path (geometry or toolpath).",
    path_index={"type": "integer", "description": "1-based index (0 = use selected)"},
    name_filter={"type": "string", "description": "Optional substring filter for attribute names"},
)
async def handle_get_path_attributes(acam, path_index=0, name_filter=None, **kw):
    return acam.get_path_attributes(path_index, name_filter)

@mcp_tool("set_path_attribute",
    "Set a user attribute on a path (geometry or toolpath).",
    attribute_name={"type": "string", "description": "Attribute name"},
    attribute_value={"type": "string", "description": "Attribute value"},
    path_index={"type": "integer", "description": "1-based index (0 = use selected)"},
)
async def handle_set_path_attribute(acam, attribute_name, attribute_value, path_index=0, **kw):
    return acam.set_path_attribute(attribute_name, attribute_value, path_index)

# ----- Operations Ordering (2) --------------------------------------------

@mcp_tool("order_operations_all",
    "Order all tool paths to match nested sheet order.")
async def handle_order_operations_all(acam, **kw):
    return acam.order_operations_all()

@mcp_tool("order_manual",
    "Reorder geometries or tool paths in a specified order by providing a list of 1-based path indices.",
    path_indices={"type": "array", "items": {"type": "integer"},
                  "description": "1-based indices of paths in the desired order"},
)
async def handle_order_manual(acam, path_indices, **kw):
    return acam.order_manual(path_indices)

# ----- Workflow (1) -------------------------------------------------------

@mcp_tool("run_workflow",
    "Run a batch of steps sequentially. Each step: {'action': 'method_name', 'params': {...}}. set_undo_point is called automatically.",
    steps={"type": "array", "items": {"type": "object"},
           "description": "List of steps, each with 'action' (method name) and optional 'params'"},
)
async def handle_run_workflow(acam, steps, **kw):
    return acam.run_workflow(steps)

# ----- View Control (3) ---------------------------------------------------

@mcp_tool("view_zoom_extents",
    "Zoom to fit all geometry on screen.")
async def handle_view_zoom_extents(acam, **kw):
    return acam.view_zoom_extents()

@mcp_tool("view_zoom_window",
    "Zoom to a rectangular window defined by two corners.",
    x1={"type": "number", "description": "First corner X"},
    y1={"type": "number", "description": "First corner Y"},
    x2={"type": "number", "description": "Second corner X"},
    y2={"type": "number", "description": "Second corner Y"},
)
async def handle_view_zoom_window(acam, x1, y1, x2, y2, **kw):
    return acam.view_zoom_window(x1, y1, x2, y2)

@mcp_tool("view_set_direction",
    "Set the 3D view direction.",
    direction={"type": "integer", "description": "0=Top, 1=Front, 2=Right, 3=Back, 4=Left, 5=Bottom, 6=SW Iso, 7=SE Iso"},
)
async def handle_view_set_direction(acam, direction, **kw):
    return acam.view_set_direction(direction)

# ----- Material (1) -------------------------------------------------------

@mcp_tool("get_material",
    "Get the current material info from the active drawing.")
async def handle_get_material(acam, **kw):
    return acam.get_material()

# ----- Undo & View (2) ----------------------------------------------------

@mcp_tool("set_undo_point",
    "Set an undo point with a descriptive name.",
    text={"type": "string", "description": "Undo point description"},
)
async def handle_set_undo_point(acam, text="AI Operation", **kw):
    return acam.set_undo_point(text)

@mcp_tool("zoom_all",
    "Zoom to extents in the active drawing.")
async def handle_zoom_all(acam, **kw):
    return acam.zoom_all()
