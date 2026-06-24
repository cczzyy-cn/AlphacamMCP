# AlphaCAM MCP Bridge — AI-Powered CAM Automation

Connect AI assistants to **AlphaCAM 2016 R1** via the [Model Context Protocol (MCP)](https://modelcontextprotocol.io/).

## Overview

```
┌─────────────┐     MCP/stdio      ┌──────────────────┐     COM      ┌─────────────────┐
│  AI Agent   │ ◄──────────────────►│  alphacam-bridge │◄────────────►│  AlphaCAM 2016  │
│  (Reasonix) │                     │  (server.py)     │   win32com   │  (Acam.exe)     │
└─────────────┘                     └──────────────────┘              └─────────────────┘
```

The bridge wraps AlphaCAM's COM Automation API (`AcamAddIns.tlb`) as **29 MCP tools** that an AI can call to read, create, modify, machine, and output NC code — all without touching the UI.

## Files

| File | Purpose |
|------|---------|
| `server.py` | MCP server — runs in stdio (Reasonix) or SSE (HTTP) mode |
| `alphacam_com.py` | Pythonic COM wrapper for AlphaCAM |
| `requirements.txt` | Python dependencies |
| `install.bat` | Windows installer (run as Admin) |

## Requirements

- **Windows** (COM is Windows-only)
- **AlphaCAM 2016 R1** installed
- **Python 3.9+** with pip
- **AlphaCAM must have been run at least once** (to register COM)

## Installation

```bash
# 1. Install Python packages
pip install mcp pywin32

# 2. Test the connection
python -c "import sys; sys.path.insert(0, '.'); from alphacam_com import AlphaCAM; a=AlphaCAM(); print(a.get_info())"
```

Or just double-click `install.bat`.

## ProgID 说明

| 模块 | ProgID |
|------|--------|
| **Router / 鉋花机**（你用的） | `aroutaps.Application` |
| Mill / 铣床 | `am5axaps.Application` |
| Lathe / 车床 | `aturhaps.Application` |
| Wire EDM / 线切割 | `awireaps.Application` |

可通过 `--progid` 参数指定模块：`python server.py --progid aroutaps.Application`

### Method 1: Reasonix MCP Integration

Add to your Reasonix `config.toml`:

```toml
[mcpServers.alphacam-bridge]
command = "python"
args = ["C:\\path\\to\\alphacam-bridge\\server.py"]
```

Then the AI can call any tool described below.

### Method 2: Standalone Test

```bash
python server.py
```

### Method 3: HTTP/SSE Mode (remote clients)

```bash
pip install uvicorn starlette
python server.py --port 8080
```

Connect MCP clients to `http://127.0.0.1:8080/sse`.

## Available MCP Tools

### Status & Info
| Tool | Description |
|------|-------------|
| `get_status` | Check AlphaCAM version, path, connection status |
| `get_drawing_info` | Active drawing: geo count, toolpaths, layers, operations |

### File Operations
| Tool | Description |
|------|-------------|
| `new_drawing` | Clear and start new drawing |
| `open_drawing` | Open `.amd` file |
| `open_dxf` | Open DXF/DWG file |
| `open_step` | Open STEP file |
| `open_stl` | Open STL file |
| `save_drawing` | Save (or Save As) active drawing |

### Geometry Creation
| Tool | Description |
|------|-------------|
| `create_rectangle` | Rectangle by 2 corners |
| `create_circle` | Circle by diameter + center |
| `create_line` | 2D line between 2 points |
| `create_polygon` | Regular polygon |
| `create_ellipse` | Ellipse |
| `create_text` | Text annotation |

### Work Plane & Layer
| Tool | Description |
|------|-------------|
| `create_workplane` | Named work plane with origin/orientation |
| `set_workplane` | Activate work plane by name |
| `create_layer` | Create or get layer, set color |

### Tool Management
| Tool | Description |
|------|-------------|
| `select_tool` | Select tool from library or show dialog |
| `get_current_tool` | Currently selected tool info |

### Machining
| Tool | Description |
|------|-------------|
| `run_machining` | Full machining operation (Rough/Finish, Pocket, Drill, Engrave) with feeds, speeds, depth etc. |
| `output_nc` | Output NC code to file |

### VBA / Add-ins
| Tool | Description |
|------|-------------|
| `run_vba_macro` | Run a VBA macro by name |
| `load_addin` | Load add-in DLL or VBA project |
| `enable_addin` | Enable/disable add-in |

### Utility
| Tool | Description |
|------|-------------|
| `set_undo_point` | Mark undo point |
| `zoom_all` | Fit view |
| `select_post` | Select post-processor |
| `list_geometries` | List all geometries |
| `list_operations` | List all operations |
| `list_toolpaths` | List all toolpaths |

### Batch
| Tool | Description |
|------|-------------|
| `run_workflow` | Multi-step batch: `[{"action":"create_rectangle", "params":{...}}, {"action":"run_machining", ...}]` |

## Example Workflow

**User says:** "Create a 100x80 rectangle, then rough-finish it with a 10mm flat endmill to depth -5mm, and output NC."

The AI calls these tools sequentially:

```
1. set_undo_point("Create rectangle + RoughFinish")
2. create_rectangle(x1=0, y1=0, x2=100, y2=80)
3. select_tool(name_or_path="$USER")  # user picks tool
4. run_machining(process_type=1, safe_rapid_level=20, final_depth=-5,
                  cut_feed=1000, spindle_speed=6000)
5. output_nc(file_path="D:\\nc\\part.nc")
6. zoom_all()
```

## Generating VBA Code with the Skill

Use the companion Reasonix skill **`alphacam-coding`** to generate VBA code:

```bash
# In Reasonix:
/alphacam-coding "Create a macro that generates a gear profile and machines it"
```

The skill will output ready-to-use VBA code using the correct API calls.

## Architecture Notes

- **Thread Safety**: AlphaCAM COM is single-threaded. The server serializes all calls.
- **Error Handling**: Each tool catches COM errors and returns descriptive messages.
- **Undo**: Use `set_undo_point` before operations for safe rollback.
- **Screen Updating**: For batch operations, set `drw.ScreenUpdating = False` via VBA for speed.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `pywin32` not found | `pip install pywin32` |
| `CLASS_E_NOTREGISTERED` | Run AlphaCAM once manually to register COM |
| `Access Denied` | Run the server as Administrator |
| `ServerNotFoundException` | Make sure AlphaCAM is installed and licensed |
| COM call hangs | AlphaCAM may have a modal dialog open — close it |
