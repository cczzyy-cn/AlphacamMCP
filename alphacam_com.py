"""
AlphaCAM COM Wrapper — provides a Pythonic interface to AlphaCAM 2016 R1 via win32com.

This layer handles:
  - Connecting to a running AlphaCAM instance or launching a new one
  - All major COM objects: Application, Drawing, MillData, MillTool, Path, etc.
  - Error handling, type conversion, safe cleanup
"""

from __future__ import annotations

import atexit
import os
import sys
from typing import Any, Optional

# ---------------------------------------------------------------------------
# Lazy imports — these only work on Windows with pywin32 installed
# ---------------------------------------------------------------------------
_win32com = None
_alphaCAMMill = None


def _ensure_com():
    global _win32com, _alphaCAMMill
    if _win32com is not None:
        return
    try:
        import win32com.client as _win32com
        from win32com.client import constants as _constants
    except ImportError:
        raise RuntimeError(
            "pywin32 is required. Install with: pip install pywin32"
        )
    # Try to import the AlphaCAM type library
    try:
        _alphaCAMMill = _win32com.GetModule(
            "{B5A0F5A0-1B5C-11D2-8F5A-006097C8E7B4}"  # AlphaCAM Mill LIBID
        )
    except Exception:
        _alphaCAMMill = None


# ---------------------------------------------------------------------------
# Constants (mirror the Acam* enums)
# ---------------------------------------------------------------------------
ACAM_LEVEL = {
    "BASIC": 1, "STANDARD": 2, "ADVANCED": 3,
    "ADVANCED3D_3AXIS": 4, "ADVANCED3D_5AXIS": 5,
    "VIEWPLUS": 6, "OEM_5AXIS": 8, "EXPRESS": 9,
}

ACAM_PROCESS = {
    "ROUGH_FINISH": 1, "CONTOUR_POCKET": 2, "MANUAL": 3,
    "SPIRAL_POCKET": 4, "LINEAR_POCKET": 5,
    "ENGRAVE": 10, "MACHINE_SURFACE": 11, "MACHINE_POLYLINE": 12,
    "DRILL": 21, "PECK": 22, "TAP": 23, "BORE": 24,
}

ACAM_DRILL_TYPE = {"DRILL": 0, "PECK": 1, "TAP": 2, "BORE": 3, "CHAMFER": 4}

ACAM_LEAD_TYPE = {
    "LINE": 0, "ARC": 1, "BOTH": 2, "NONE": 3, "NO_CHANGE": 4,
}

ACAM_LEAD_SIDE = {"LEFT": 0, "CENTER": 1, "RIGHT": 2}

ACAM_TOOL_TYPE = {
    "SQUARE": 0, "BULL": 1, "BALL": 2, "DRILL": 3, "TAP": 4,
    "USER": 5, "WHEEL": 6, "LOLLIPOP": 7,
}

ACAM_TOOL_SIDE = {"INSIDE": 0, "OUTSIDE": 1, "ON": 2}

ACAM_OUTPUT = {"FILE": 0, "SCREEN": 1, "PRINTER": 2}

ACAM_MENU = {
    "ADDINS": 1, "FILE": 2, "FILE_OPEN": 3, "FILE_SAVE": 4,
    "FILE_INPUT": 5, "FILE_OUTPUT": 6, "FILE_POST": 7,
    "FILE_CONFIGURE": 8, "EDIT": 9, "VIEW": 10, "GEO": 11,
    "3D": 12, "UTILS": 13, "CAD": 14, "MACHINE": 15,
    "HELP": 16, "NEW": 17, "SPECIALGEO": 18, "SPECIALFUN": 19,
    "MACHINE_SETUP": 20, "MACHINE_OPS": 21, "MACHINE_3D": 22,
    "MACHINE_EDIT": 23,
}


# ---------------------------------------------------------------------------
# COM wrapper exceptions
# ---------------------------------------------------------------------------
class AlphaCAMError(RuntimeError):
    """Generic AlphaCAM automation error."""


class AlphaCAMNotRunning(AlphaCAMError):
    """AlphaCAM is not running or could not be started."""


# ---------------------------------------------------------------------------
# Main wrapper
# ---------------------------------------------------------------------------
class AlphaCAM:
    """Wraps the AlphaCAM Application COM object."""

    def __init__(self, prog_id: str = "aroutaps.Application",
                 visible: bool = True):
        _ensure_com()
        self._prog_id = prog_id
        self._app: Any = None
        self._visible = visible
        self._connect()
        atexit.register(self._cleanup)

    # ---- connect / disconnect --------------------------------------------

    def _connect(self):
        """
        Connect to AlphaCAM via COM automation.
        
        Note: AlphaCAM must be started via COM (CreateObject) for the MCP bridge
        to connect to it. Manually launched instances are not registered in COM's 
        Running Object Table.
        
        Strategy:
          1. Try GetActiveObject (attaches to an existing COM-launched instance)
          2. Fall back to CreateObject/Dispatch (starts AlphaCAM fresh via COM)
        """
        try:
            self._app = _win32com.GetActiveObject(self._prog_id)
        except Exception:
            try:
                self._app = _win32com.Dispatch(self._prog_id)
            except Exception as exc:
                raise AlphaCAMNotRunning(
                    f"Cannot connect to AlphaCAM ({self._prog_id}): {exc}"
                ) from exc
        self._app.Visible = self._visible

    def _cleanup(self):
        """Release the COM object."""
        if self._app is not None:
            try:
                self._app.Visible = True
            except Exception:
                pass
            self._app = None

    def restart(self):
        """Disconnect and reconnect."""
        self._cleanup()
        self._connect()

    @property
    def app(self) -> Any:
        """The raw COM Application object."""
        return self._app

    @property
    def is_connected(self) -> bool:
        try:
            _ = self._app.Name
            return True
        except Exception:
            return False

    # ---- info / status ---------------------------------------------------

    def get_info(self) -> dict:
        """Return basic info about the AlphaCAM instance."""
        return {
            "name": self._safe_get(self._app, "Name", ""),
            "full_name": self._safe_get(self._app, "FullName", ""),
            "path": self._safe_get(self._app, "Path", ""),
            "version": str(self._safe_get(self._app, "AlphacamVersion", "")),
            "api_version": self._safe_get(self._app, "ApiVersion", 0),
            "program_letter": chr(self._safe_get(self._app, "ProgramLetter", 77)),
            "program_level": self._safe_get(self._app, "ProgramLevel", 0),
            "licomdat_path": self._safe_get(self._app, "LicomdatPath", ""),
            "licomdir_path": self._safe_get(self._app, "LicomdirPath", ""),
            "post_file_name": self._safe_get(self._app, "PostFileName", ""),
            "visible": self._safe_get(self._app, "Visible", False),
            "is_connected": self.is_connected,
        }

    # ---- drawing management ----------------------------------------------

    @property
    def active_drawing(self) -> Any:
        """Get the active Drawing object (or None)."""
        try:
            return self._app.ActiveDrawing
        except Exception:
            return None

    def new_drawing(self):
        """Clear everything and start a new drawing."""
        self._app.New()

    def open_drawing(self, file_path: str) -> dict:
        """Open an AlphaCAM drawing file (.amd etc.)."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        drw = self._app.OpenDrawing(file_path)
        return self._drawing_info(drw)

    def open_dxf(self, file_path: str, clear: bool = True) -> dict:
        """Open a DXF/DWG file."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        drw = self._app.OpenDxfFile(file_path, clear)
        return self._drawing_info(drw)

    def open_step(self, file_path: str, clear: bool = True) -> dict:
        """Open a STEP file."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        drw = self._app.OpenStepFile(file_path, clear)
        return self._drawing_info(drw)

    def open_stl(self, file_path: str, clear: bool = True) -> dict:
        """Open an STL file."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        drw = self._app.OpenStlFile(file_path, clear)
        return self._drawing_info(drw)

    def open_external(self, file_path: str, clear: bool = True) -> dict:
        """Open via external import add-in (Parasolid, Inventor, etc.)."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        drw = self._app.OpenExternalFile(file_path, clear)
        return self._drawing_info(drw)

    def save_drawing(self, file_path: Optional[str] = None):
        """Save the active drawing."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        if file_path:
            drw.SaveAs(file_path)
        else:
            drw.Save()

    def close_drawing(self):
        """Close without saving — starts new empty drawing."""
        self._app.New()

    # ---- drawing info ----------------------------------------------------

    def _drawing_info(self, drw: Any) -> dict:
        """Extract info from a Drawing object."""
        if drw is None:
            return {"error": "No active drawing"}
        try:
            geos = self._safe_get(drw, "Geometries", None)
            tps = self._safe_get(drw, "ToolPaths", None)
            ops = self._safe_get(drw, "Operations", None)
            layers = self._safe_get(drw, "Layers", None)
            surfs = self._safe_get(drw, "Surfaces", None)
            return {
                "name": self._safe_get(drw, "Name", ""),
                "full_name": self._safe_get(drw, "FullName", ""),
                "modified": bool(self._safe_get(drw, "Modified", False)),
                "geo_count": self._safe_get(drw, "GetGeoCount", 0),
                "toolpath_count": self._safe_get(drw, "GetToolPathCount", 0),
                "layer_count": self._safe_get(drw, "GetLayerCount", 0),
                "operations_count": (
                    self._safe_get(ops, "Count", 0) if ops else 0
                ),
            }
        except Exception as exc:
            return {"error": str(exc)}

    def get_drawing_info(self) -> dict:
        """Info about the active drawing."""
        return self._drawing_info(self.active_drawing)

    def list_geometries(self) -> list[dict]:
        """List all geometries in the active drawing."""
        drw = self.active_drawing
        if drw is None:
            return []
        result = []
        try:
            count = drw.GetGeoCount()
            first = drw.GetFirstGeo()
            last = drw.GetLastGeo()
            if first is not None:
                result.append({
                    "index": 1,
                    "type": self._path_type_name(first),
                    "length": round(self._safe_get(first, "Length", 0.0), 3),
                    "closed": bool(self._safe_get(first, "Closed", False)),
                    "selected": bool(self._safe_get(first, "Selected", False)),
                })
            result.append({
                "total_count": count,
                "has_multiple": count > 1,
            })
        except Exception as exc:
            result.append({"error": str(exc)})
        return result

    def list_toolpaths(self) -> list[dict]:
        """List all toolpaths."""
        drw = self.active_drawing
        if drw is None:
            return []
        result = []
        try:
            count = drw.GetToolPathCount()
            first = drw.GetFirstToolPath()
            last = drw.GetLastToolPath()
            if first is not None:
                result.append({
                    "index": 1,
                    "length": round(self._safe_get(first, "Length", 0.0), 3),
                })
            result.append({
                "total_count": count,
                "has_multiple": count > 1,
            })
        except Exception as exc:
            result.append({"error": str(exc)})
        return result

    def list_operations(self) -> list[dict]:
        """List all operations."""
        drw = self.active_drawing
        if drw is None:
            return []
        result = []
        ops = self._safe_get(drw, "Operations", None)
        if ops is None:
            return []
        for i in range(1, ops.Count + 1):
            op = ops.Item(i)
            try:
                subs = self._safe_get(op, "SubOperations", None)
                sub_count = self._safe_get(subs, "Count", 0) if subs else 0
                result.append({
                    "number": self._safe_get(op, "Number", 0),
                    "visible": bool(self._safe_get(op, "Visible", True)),
                    "sub_operation_count": sub_count,
                })
            except Exception:
                result.append({"number": i})
        return result

    # ---- geometry creation -----------------------------------------------

    def create_rectangle(self, x1: float, y1: float,
                         x2: float, y2: float) -> dict:
        """Create a rectangle."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = drw.CreateRectangle(x1, y1, x2, y2)
        return {"path_type": self._path_type_name(p), "length": p.Length}

    def create_circle(self, diameter: float, xc: float, yc: float) -> dict:
        """Create a circle by diameter and center."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = drw.CreateCircle(diameter, xc, yc)
        return {"path_type": self._path_type_name(p), "length": p.Length}

    def create_circle_3pts(self, x1, y1, x2, y2, x3, y3) -> dict:
        """Create a circle by 3 points."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = drw.CreateCircle2(x1, y1, x2, y2, x3, y3)
        return {"path_type": self._path_type_name(p), "length": p.Length}

    def create_line(self, x1, y1, x2, y2) -> dict:
        """Create a 2D line."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = drw.Create2DLine(x1, y1, x2, y2)
        return {"path_type": self._path_type_name(p), "length": p.Length}

    def create_polygon(self, xc, yc, radius, sides,
                       start_angle=0) -> dict:
        """Create a regular polygon."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = drw.CreatePolygon(xc, yc, radius, sides, start_angle)
        return {"path_type": self._path_type_name(p), "length": p.Length}

    def create_ellipse(self, xc, yc, maj_ax, min_ax, angle=0) -> dict:
        """Create an ellipse."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = drw.CreateEllipse(xc, yc, maj_ax, min_ax, angle)
        return {"path_type": self._path_type_name(p), "length": p.Length}

    def create_text(self, text: str, height: float, x: float, y: float,
                    font: str = "Arial", align: int = 0,
                    angle: float = 0) -> dict:
        """Create a text object."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        t = drw.CreateText(height, angle, text, x, y, font, align)
        return {"text": text, "height": height, "x": x, "y": y}

    # ---- workplane & layer -----------------------------------------------

    def create_workplane(self, name: str, x: float = 0, y: float = 0,
                         z: float = 0, i: float = 0, j: float = 0,
                         k: float = 1, x2: float = 1, y2: float = 0) -> dict:
        """Create a work plane."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        wp = drw.CreateWorkPlane(name, x, y, z, i, j, k, x2, y2)
        return {"name": wp.Name, "id": wp.Id}

    def set_workplane(self, name: str) -> bool:
        """Set the active work plane by name."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        wp = drw.GetWorkPlane(name)
        if wp is None:
            raise AlphaCAMError(f"Work plane '{name}' not found")
        drw.SetWorkPlane(wp)
        return True

    def create_layer(self, name: str, color: Optional[int] = None) -> dict:
        """Create or get a layer."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        lyr = drw.CreateLayer(name)
        if color is not None:
            lyr.ColorRGB = color
        return {"name": lyr.Name, "active": lyr.Active,
                "visible": lyr.Visible, "locked": lyr.Lock}

    def list_layers(self) -> list[dict]:
        """List all layers."""
        drw = self.active_drawing
        if drw is None:
            return []
        result = []
        try:
            count = drw.GetLayerCount()
            for i in range(1, count + 1):
                lyr = drw.GetLayer(i)
                result.append({
                    "name": lyr.Name,
                    "active": lyr.Active,
                    "visible": lyr.Visible,
                    "locked": lyr.Lock,
                })
        except Exception as exc:
            result.append({"error": str(exc)})
        return result

    # ---- tool selection --------------------------------------------------

    def delete_selected(self):
        """Delete all selected geometries in the active drawing."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        count = 0
        p = drw.GetFirstGeo()
        while p is not None:
            try:
                if p.Selected:
                    p.Erase()
                    count += 1
            except Exception:
                pass
            try:
                p = drw.GetLastGeo()
            except Exception:
                break
        return {"deleted_count": count}

    def delete_all_geometries(self):
        """Delete all geometries (keeps toolpaths, layers, etc.)."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        drw.Clear(True, False, False, False, False, False, False, False)
        return {"status": "ok", "action": "all geometries cleared"}

    def trim_with_boundary(self, boundary_index: int = 0,
                           line_indices: Optional[list[int]] = None) -> dict:
        """
        Trim lines by a boundary (closed path). Uses TrimWithCuttingGeos to
        truncate lines at boundary edges and remove external portions.

        Args:
            boundary_index: 1-based index of the boundary geometry (0=use selected)
            line_indices: 1-based indices of lines to trim (None = all intersecting)
        """
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")

        first = drw.GetFirstGeo()
        if first is None:
            raise AlphaCAMError("No geometries in drawing")

        boundary = self._find_geo_by_index_or_selected(boundary_index)
        if boundary is None:
            raise AlphaCAMError("Boundary geometry not found or not selected")

        if not boundary.Closed:
            raise AlphaCAMError("Boundary geometry must be a closed path")

        # Collect lines to trim
        lines_to_trim = self._collect_lines_to_trim(boundary, line_indices)
        if not lines_to_trim:
            return {"boundary_found": True, "total_lines_trimmed": 0,
                    "message": "No intersecting lines found"}

        total_trims = 0
        processed = 0
        errors = []

        for line in lines_to_trim:
            try:
                if not line.TestIntersectPath(boundary, 0, 0):
                    processed += 1
                    continue

                # Select boundary as cutting tool
                drw.SetGeosSelected(False)
                boundary.Selected = True

                # Get both endpoint coordinates
                first_elem = line.GetFirstElem()
                last_elem = line.GetLastElem()
                sx, sy = first_elem.StartXG, first_elem.StartYG
                ex, ey = last_elem.EndXG, last_elem.EndYG

                # Trim each external endpoint (TrimWithCuttingGeos deletes
                # the portion containing the specified point)
                if not boundary.IsPointInside(sx, sy):
                    line.TrimWithCuttingGeos(sx, sy)
                    total_trims += 1
                    drw.SetGeosSelected(False)
                    boundary.Selected = True

                if not boundary.IsPointInside(ex, ey):
                    line.TrimWithCuttingGeos(ex, ey)
                    total_trims += 1

                processed += 1

            except Exception as exc:
                errors.append(str(exc))
                processed += 1

        drw.SetGeosSelected(False)
        drw.Redraw()

        return {
            "boundary_found": True,
            "total_lines_processed": processed,
            "total_trims_applied": total_trims,
            "errors": errors if errors else None,
        }

    def select_tool(self, name_or_path: str = "$USER") -> dict:
        """Select a tool. Use '$USER' for the tool selection dialog."""
        try:
            tool = self._app.SelectTool(name_or_path)
            if tool is None:
                return {"error": "No tool selected", "name": name_or_path}
            return {
                "name": self._safe_get(tool, "Name", ""),
                "diameter": self._safe_get(tool, "Diameter", 0.0),
                "type": self._safe_get(tool, "Type", -1),
                "number": self._safe_get(tool, "Number", 0),
                "corner_radius": self._safe_get(tool, "CornerRadius", 0.0),
                "length": self._safe_get(tool, "Length", 0.0),
                "spindle_speed": self._safe_get(tool, "SpindleSpeed", 0.0),
                "num_teeth": self._safe_get(tool, "NumberOfTeeth", 0),
            }
        except Exception as exc:
            raise AlphaCAMError(f"Tool selection failed: {exc}") from exc

    def get_current_tool(self) -> dict:
        """Get the currently selected tool."""
        tool = self._app.GetCurrentTool()
        if tool is None:
            return {"error": "No tool selected"}
        return {
            "name": self._safe_get(tool, "Name", ""),
            "diameter": self._safe_get(tool, "Diameter", 0.0),
            "number": self._safe_get(tool, "Number", 0),
        }

    # ---- machining -------------------------------------------------------

    def create_mill_data(self, params: dict) -> Any:
        """Create a MillData object and set properties from a dict."""
        md = self._app.CreateMillData()
        prop_map = {
            "safe_rapid_level": "SafeRapidLevel",
            "rapid_down_to": "RapidDownTo",
            "final_depth": "FinalDepth",
            "material_top": "MaterialTop",
            "depth_of_cut": "DepthOfCut",
            "max_depth_per_cut": "MaxDepthPerCut",
            "cut_feed": "CutFeed",
            "down_feed": "DownFeed",
            "spindle_speed": "SpindleSpeed",
            "stock": "Stock",
            "stock_xy": "StockXY",
            "stock_z": "StockZ",
            "number_of_cuts": "NumberOfCuts",
            "bidirectional": "Bidirectional",
            "width_of_cut": "WidthOfCut",
            "process_type": "ProcessType2",
            "drill_type": "DrillType",
            "bottom_of_hole": "BottomOfHole",
            "peck_distance": "PeckDistance",
            "dwell_time": "DwellTime",
            "thread_pitch": "ThreadPitch",
            "coolant": "Coolant",
            "mc_comp": "McComp",
            "chord_error": "ChordError",
            "tool_number": "ToolNumber",
            "offset_number": "OffsetNumber",
            "op_no": "OpNo",
            "sides": "Sides",
            "side_angle": "SideAngle",
            "pocket_type": "PocketType",
        }
        for py_name, acam_name in prop_map.items():
            if py_name in params:
                try:
                    setattr(md, acam_name, params[py_name])
                except Exception:
                    pass  # skip invalid properties
        return md

    def run_machining(self, mill_data_params: dict,
                      select_geo: bool = True) -> dict:
        """Run a machining operation with the given parameters."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")

        md = self.create_mill_data(mill_data_params)

        process = mill_data_params.get("process_type", 1)
        result = None

        try:
            if process in (1, 2, 3, 4, 5):  # RoughFinish, Pocket, etc.
                result = md.RoughFinish()
            elif process == 10:  # Engrave
                result = md.Engrave()
            elif process in (21, 22, 23, 24):  # Drill
                result = md.DrillTap()
            else:
                result = md.RoughFinish()
        except Exception as exc:
            raise AlphaCAMError(f"Machining failed: {exc}") from exc

        paths_created = 0
        if result is not None:
            try:
                paths_created = result.Count
            except Exception:
                paths_created = 1

        drw.ZoomAll()
        return {
            "paths_created": paths_created,
            "process_type": process,
            "parameters": mill_data_params,
        }

    def output_nc(self, file_path: str,
                  output_to: int = 0,
                  visible_only: bool = True) -> dict:
        """Output NC code to file."""
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        # Ensure directory exists
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        try:
            drw.OutputNC(file_path, output_to, visible_only)
        except Exception as exc:
            raise AlphaCAMError(f"NC output failed: {exc}") from exc
        return {
            "file_path": file_path,
            "output_type": output_to,
            "success": os.path.exists(file_path),
        }

    # ---- VBA / automation ------------------------------------------------

    def run_vba_macro(self, macro_name: str,
                      params: Optional[list] = None) -> Any:
        """Run a VBA macro by name with optional parameters."""
        args = [macro_name]
        if params:
            args.extend(params[:8])  # max 8 params
        try:
            result = self._app.Run(*args)
            return {"result": str(result) if result is not None else None}
        except Exception as exc:
            raise AlphaCAMError(f"VBA macro failed: {exc}") from exc

    def load_addin(self, file_name: str) -> dict:
        """Load an add-in DLL or VBA project file."""
        try:
            self._app.LoadAddIn(file_name)
            return {"loaded": file_name}
        except Exception as exc:
            raise AlphaCAMError(f"Add-in load failed: {exc}") from exc

    def enable_addin(self, name: str, enable: bool = True) -> dict:
        """Enable or disable an add-in."""
        try:
            self._app.EnableAddin(name, enable)
            return {"addin": name, "enabled": enable}
        except Exception as exc:
            raise AlphaCAMError(f"EnableAddin failed: {exc}") from exc

    # ---- utilities -------------------------------------------------------

    def shell_and_wait(self, command: str) -> dict:
        """Run an external command and wait for it to finish."""
        self._app.ShellAndWait(command, True)
        return {"command": command}

    def set_undo_point(self, text: str = "AI Operation"):
        """Set an undo point with the given name."""
        # Note: SetUndoPoint COM method has a known parameter binding issue
        # in win32com. We use SetUndoCommandName + SetUndoPoint() instead.
        try:
            self._app.SetUndoCommandName(text)
            self._app.SetUndoPoint()
        except Exception:
            # Fallback: try direct call
            try:
                self._app.SetUndoPoint()
            except Exception:
                pass
        return {"undo_point": text}

    def zoom_all(self):
        """Zoom extents."""
        drw = self.active_drawing
        if drw:
            drw.ZoomAll()

    def select_post(self, post_name: str) -> dict:
        """Select a post-processor by name."""
        try:
            self._app.SelectPost(post_name)
            return {"post": post_name}
        except Exception as exc:
            raise AlphaCAMError(f"SelectPost failed: {exc}") from exc



    # ---- RevNest operations (screen locking, nesting, path ops) -------

    def list_addins(self) -> list:
        """List all loaded add-ins."""
        result = []
        try:
            for ai in self._app.AddIns:
                result.append({
                    "name": ai.Name,
                    "connected": bool(ai.Connect),
                    "description": ai.Description if hasattr(ai, 'Description') else "",
                })
        except Exception:
            pass
        return result

    def lock_acam(self) -> dict:
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        try:
            self._app.Frame.ProjectBarUpdating = False
            drw.ScreenUpdating = False
            return {"status": "locked"}
        except Exception as exc:
            raise AlphaCAMError(f"Lock failed: {exc}") from exc

    def unlock_acam(self, zoom_all: bool = False) -> dict:
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        try:
            self._app.Frame.ProjectBarUpdating = True
            if zoom_all:
                drw.ZoomAll()
            drw.ScreenUpdating = True
            return {"status": "unlocked", "zoom_all": zoom_all}
        except Exception as exc:
            raise AlphaCAMError(f"Unlock failed: {exc}") from exc

    def has_nesting(self) -> dict:
        drw = self.active_drawing
        if drw is None:
            return {"has_nesting": False, "error": "No active drawing"}
        try:
            ni = drw.GetNestInformation()
            if ni is None:
                return {"has_nesting": False}
            sheet_count = ni.Sheets.Count
            return {"has_nesting": sheet_count > 0, "sheet_count": sheet_count}
        except Exception:
            return {"has_nesting": False}

    def get_nesting_info(self) -> dict:
        drw = self.active_drawing
        if drw is None:
            return {"error": "No active drawing"}
        try:
            ni = drw.GetNestInformation()
            if ni is None:
                return {"has_nesting": False}
            sheets_info = []
            for sh in ni.Sheets:
                geo = sh.Geometry
                sheet_data = {
                    "name": self._safe_get(geo, "Name", ""),
                    "min_x": self._safe_get(geo, "MinXL", 0.0),
                    "min_y": self._safe_get(geo, "MinYL", 0.0),
                    "max_x": self._safe_get(geo, "MaxXL", 0.0),
                    "max_y": self._safe_get(geo, "MaxYL", 0.0),
                }
                parts_list = []
                for inst in sh.Parts:
                    pgeo = inst.Geometry
                    parts_list.append({
                        "name": self._safe_get(inst, "Name", ""),
                        "pos_x": self._safe_get(pgeo, "MinXL", 0.0),
                        "pos_y": self._safe_get(pgeo, "MinYL", 0.0),
                        "width": self._safe_get(pgeo, "MaxXL", 0.0) - self._safe_get(pgeo, "MinXL", 0.0),
                        "height": self._safe_get(pgeo, "MaxYL", 0.0) - self._safe_get(pgeo, "MinYL", 0.0),
                    })
                sheet_data["parts"] = parts_list
                sheet_data["part_count"] = len(parts_list)
                sheets_info.append(sheet_data)
            return {
                "has_nesting": True,
                "sheet_count": ni.Sheets.Count,
                "sheets": sheets_info,
            }
        except Exception as exc:
            return {"error": str(exc)}

    def get_sheet_extents(self) -> dict:
        drw = self.active_drawing
        if drw is None:
            return {"error": "No active drawing"}
        try:
            ni = drw.GetNestInformation()
            if ni is None or ni.Sheets.Count == 0:
                return {"has_sheets": False}
            min_x = min_y = 1e20
            max_x = max_y = -1e20
            for sh in ni.Sheets:
                geo = sh.Geometry
                mx1 = self._safe_get(geo, "MinXL", 1e20)
                my1 = self._safe_get(geo, "MinYL", 1e20)
                mx2 = self._safe_get(geo, "MaxXL", -1e20)
                my2 = self._safe_get(geo, "MaxYL", -1e20)
                if mx1 < min_x: min_x = mx1
                if my1 < min_y: min_y = my1
                if mx2 > max_x: max_x = mx2
                if my2 > max_y: max_y = my2
            return {
                "has_sheets": True,
                "sheet_count": ni.Sheets.Count,
                "min_x": min_x, "min_y": min_y,
                "max_x": max_x, "max_y": max_y,
                "width": max_x - min_x,
                "height": max_y - min_y,
            }
        except Exception as exc:
            return {"error": str(exc)}

    def mirror_path(self, x1: float, y1: float, x2: float, y2: float, path_index: int = 0) -> dict:
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = self._find_geo_by_index_or_selected(path_index)
        if p is None:
            raise AlphaCAMError("No path found")
        try:
            p.MirrorL(x1, y1, x2, y2)
            return {"status": "mirrored", "mirror_line": {"x1": x1, "y1": y1, "x2": x2, "y2": y2}}
        except Exception as exc:
            raise AlphaCAMError(f"MirrorL failed: {exc}") from exc

    def copy_temporary_store(self, path_index: int = 0, mirror: dict = None) -> dict:
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = self._find_geo_by_index_or_selected(path_index)
        if p is None:
            raise AlphaCAMError("No path found")
        try:
            pcopy = p.CopyTemporary()
            if mirror:
                pcopy.MirrorL(mirror.get("x1",0), mirror.get("y1",0), mirror.get("x2",0), mirror.get("y2",0))
            pcopy.StoreTemporary()
            return {
                "status": "stored",
                "original_length": round(self._safe_get(p, "Length", 0.0), 3),
                "copy_length": round(self._safe_get(pcopy, "Length", 0.0), 3),
                "applied_mirror": mirror is not None,
            }
        except Exception as exc:
            raise AlphaCAMError(f"CopyTemporary/StoreTemporary failed: {exc}") from exc

    def get_path_attributes(self, path_index: int = 0, name_filter: str = None) -> dict:
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = self._find_geo_by_index_or_selected(path_index)
        if p is None:
            raise AlphaCAMError("No path found")
        try:
            count = p.GetAttributeCount()
            attrs = {}
            for i in range(1, count + 1):
                aname = p.GetAttributeName(i)
                if name_filter and name_filter not in aname:
                    continue
                try:
                    attrs[aname] = str(p.Attribute(aname))
                except Exception:
                    attrs[aname] = "<error reading>"
            return {"path_index": path_index if path_index else "selected", "attribute_count": count, "returned_count": len(attrs), "attributes": attrs}
        except Exception as exc:
            raise AlphaCAMError(f"Get attributes failed: {exc}") from exc

    def set_path_attribute(self, attribute_name: str, attribute_value, path_index: int = 0) -> dict:
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = self._find_geo_by_index_or_selected(path_index)
        if p is None:
            raise AlphaCAMError("No path found")
        try:
            # COM property put with parameter - use Invoke with DISPATCH_PROPERTYPUT
            import pythoncom
            dispid = p._disp_.GetIDsOfNames("Attribute")[0]
            p._oleobj_.Invoke(dispid, pythoncom.IID_NULL, 0, pythoncom.DISPATCH_PROPERTYPUT, attribute_name, attribute_value)
            return {"status": "set", "attribute_name": attribute_name, "attribute_value": str(attribute_value)}
        except Exception as exc:
            raise AlphaCAMError(f"Set attribute failed: {exc}") from exc

    def offset_path(self, distance: float, side: int = 1, path_index: int = 0, delete_original: bool = False) -> dict:
        drw = self.active_drawing
        if drw is None:
            raise AlphaCAMError("No active drawing")
        p = self._find_geo_by_index_or_selected(path_index)
        if p is None:
            raise AlphaCAMError("No path found")
        try:
            pths_ret = p.Offset(distance, side)
            if pths_ret is None or pths_ret.Count == 0:
                raise AlphaCAMError("Offset returned no paths")
            first = pths_ret.Item(1)
            result = {"status": "offset", "original_length": round(self._safe_get(p, "Length", 0.0), 3), "offset_length": round(self._safe_get(first, "Length", 0.0), 3), "offset_count": pths_ret.Count}
            if delete_original:
                p.Delete()
                result["original_deleted"] = True
            return result
        except Exception as exc:
            raise AlphaCAMError(f"Offset failed: {exc}") from exc

    def get_all_geometries(self, include_attributes: bool = False) -> list:
        drw = self.active_drawing
        if drw is None:
            return []
        result = []
        try:
            idx = 1
            p = drw.GetFirstGeo()
            while p is not None:
                entry = {
                    "index": idx,
                    "type": "Closed" if self._safe_get(p, "Closed", False) else "Open",
                    "length": round(self._safe_get(p, "Length", 0.0), 3),
                    "min_x": self._safe_get(p, "MinXL", 0.0),
                    "min_y": self._safe_get(p, "MinYL", 0.0),
                    "max_x": self._safe_get(p, "MaxXL", 0.0),
                    "max_y": self._safe_get(p, "MaxYL", 0.0),
                    "selected": bool(self._safe_get(p, "Selected", False)),
                    "sheet": bool(self._safe_get(p, "Sheet", False)),
                    "dimension": bool(self._safe_get(p, "Dimension", False)),
                    "disabled": bool(self._safe_get(p, "Disabled", False)),
                    "visible": bool(self._safe_get(p, "Visible", True)),
                    "name": self._safe_get(p, "Name", ""),
                    "group": self._safe_get(p, "Group", 0),
                }
                if include_attributes:
                    try:
                        acount = p.GetAttributeCount()
                        attrs = {}
                        for ai in range(1, acount + 1):
                            aname = p.GetAttributeName(ai)
                            try:
                                attrs[aname] = str(p.Attribute(aname))
                            except Exception:
                                attrs[aname] = "<error>"
                        entry["attributes"] = attrs
                        entry["attribute_count"] = acount
                    except Exception:
                        entry["attributes"] = {}
                        entry["attribute_count"] = 0
                result.append(entry)
                idx += 1
                try:
                    p = p.GetNext()
                except Exception:
                    break
            result.append({"total_count": len(result)})
        except Exception as exc:
            result.append({"error": str(exc)})
        return result

    # ---- helpers ---------------------------------------------------------

    @staticmethod
    def _safe_get(obj: Any, attr: str, default: Any = None) -> Any:
        """Safely get an attribute/method from a COM object."""
        if obj is None:
            return default
        try:
            val = getattr(obj, attr)
            if callable(val):
                return val()
            return val
        except Exception:
            return default

    @staticmethod
    def _path_type_name(p: Any) -> str:
        """Return a human-readable path type name."""
        try:
            if p.Closed:
                return "Closed"
            return "Open"
        except Exception:
            return "Unknown"

    def _find_geo_by_index_or_selected(self, index: int = 0):
        """Find a geometry by 1-based index (0 = use selected)."""
        drw = self.active_drawing
        if drw is None:
            return None
        if index == 0:
            p = drw.GetFirstGeo()
            while p is not None:
                try:
                    if p.Selected:
                        return p
                    p = p.GetNext()
                except Exception:
                    break
            return None
        idx = 1
        p = drw.GetFirstGeo()
        while p is not None:
            if idx == index:
                return p
            idx += 1
            try:
                p = p.GetNext()
            except Exception:
                break
        return None

    def _collect_lines_to_trim(self, boundary, line_indices=None):
        """Collect lines to trim. line_indices=None = all intersecting."""
        drw = self.active_drawing
        if drw is None:
            return []
        result = []
        # Get boundary identity marker
        boundary_len = boundary.Length
        boundary_closed = boundary.Closed
        boundary_name = self._safe_get(boundary, "Name", "")

        if line_indices:
            idx = 1
            p = drw.GetFirstGeo()
            while p is not None:
                if idx in line_indices:
                    # Check it's not the boundary
                    p_len = self._safe_get(p, "Length", -1)
                    if abs(p_len - boundary_len) > 0.001 or not p.Closed:
                        result.append(p)
                idx += 1
                try:
                    p = p.GetNext()
                except Exception:
                    break
        else:
            p = drw.GetFirstGeo()
            while p is not None:
                # Skip boundary by comparing Length + Closed status
                p_len = self._safe_get(p, "Length", -1)
                is_boundary = (abs(p_len - boundary_len) < 0.001
                               and p.Closed == boundary_closed)
                if not is_boundary:
                    try:
                        if p.TestIntersectPath(boundary, 0, 0):
                            result.append(p)
                    except Exception:
                        pass
                try:
                    p = p.GetNext()
                except Exception:
                    break
        return result

    # ---- batch / convenience ---------------------------------------------

    def run_workflow(self, steps: list[dict]) -> list[dict]:
        """
        Run a batch of steps sequentially.
        Each step: {"action": "...", "params": {...}}
        """
        results = []
        for step in steps:
            action = step.get("action", "")
            params = step.get("params", {})
            try:
                handler = getattr(self, action, None)
                if handler is None:
                    results.append({"action": action,
                                    "error": f"Unknown action: {action}"})
                    continue
                self.set_undo_point(f"AI: {action}")
                result = handler(**params)
                results.append({"action": action, "result": result})
            except Exception as exc:
                results.append({"action": action, "error": str(exc)})
                break
        return results
