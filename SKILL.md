---
name: alphacam-bridge
description: Guide for using the alphacam-bridge MCP plugin to directly operate AlphaCAM 2016 R1 in real-time.
---

# AlphaCAM MCP Bridge — AI to operate AlphaCAM in real-time

You have **29 MCP tools** available via the `alphacam-bridge` plugin. Use them to directly control AlphaCAM 2016 R1.

## Available Tools

### Status & Info
- **`get_status`** — Check connection, version, paths
- **`get_drawing_info`** — Active drawing details (geos, TPs, layers, ops)
- **`list_geometries`** / **`list_operations`** / **`list_toolpaths`** — List entities

### File
- **`new_drawing`** / **`open_drawing`** / **`open_dxf`** / **`open_step`** / **`open_stl`**
- **`save_drawing`**

### Geometry Creation
- **`create_rectangle`** / **`create_circle`** / **`create_line`** / **`create_polygon`** / **`create_ellipse`** / **`create_text`**

### Workplane & Layer
- **`create_workplane`** / **`set_workplane`** / **`create_layer`**

### Tool
- **`select_tool`** / **`get_current_tool`**

### Machining
- **`run_machining`** — Full process control (feeds, speeds, depth, process type)
- **`output_nc`** — Write NC code to file

### VBA & Add-ins
- **`run_vba_macro`** — Call any VBA macro
- **`load_addin`** / **`enable_addin`**

### Utilities
- **`set_undo_point`** / **`zoom_all`** / **`select_post`**

### Batch
- **`run_workflow`** — Chain multiple steps in one call

## Process Type Constants
- 1 = Rough/Finish
- 2 = Contour Pocket
- 10 = Engrave
- 21 = Drill / 22 = Peck / 23 = Tap / 24 = Bore

## When to Use MCP Tools vs. VBA Code

| Scenario | Use |
|----------|-----|
| Quick geometry creation | MCP tools |
| Standard machining | MCP tools |
| Complex batch operations | `run_workflow` |
| Custom logic / loops / conditions | Generate VBA via `alphacam-coding` skill |
| Custom menu items / add-ins | Generate VBA via `alphacam-coding` skill |
| Event handlers (BeforeSave, AfterOutput) | Generate VBA via `alphacam-coding` skill |
| Interactive user selection | MCP tools (`select_tool`, etc.) |
| Multi-step automated workflow | MCP `run_workflow`

## TFS#80910: WorkPlane Check

When inserting paths from a temp drawing, always check for workplanes:

```bash
# After OpenTempDrawing, check WorkPlanes count
# If > 0, the drawing contains workplanes and cannot be inserted
```

This is checked via the `Drawing.WorkPlanes` property.
