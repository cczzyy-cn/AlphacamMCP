# AlphaCAM MCP 桥接器 — AI 驱动的 CAM 自动化

通过 [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) 让 AI 助手直接操作 **AlphaCAM 2016 R1**。

## 概述

本桥接器将 AlphaCAM 2016 R1 的 COM API 封装为 MCP 工具，使 AI 能够实时：
- 创建/读取/修改几何图形（线、圆、矩形、多边形、椭圆）
- 执行加工操作（粗精加工、轮廓铣槽、雕刻、钻孔等）
- 管理刀具和刀具路径
- 处理排版嵌套（Nesting）
- 运行 VBA 宏
- 输出 NC 代码

## 文件说明

| 文件 | 说明 |
|---|---|
| `server.py` | MCP 桥接器主程序（Python），通过 STDIO/SSE 协议与 AI 通信 |
| `alphacam_com.py` | AlphaCAM COM 自动化封装层（含 VBA 模块管理、自动重连） |
| `DOCUMENTATION_INDEX.md` | AlphaCAM 全部 33 个 .chm 文档的索引目录（含分组和转换状态） |
| `chm/` | .chm 文档目录（含指向安装目录的符号链接 + 已转换的 _html 子目录） |
| `CCC功能/` | VBA 插件合集目录（依边界裁剪、全排版刀具偏移、排版刀具排序） |
| `CDM功能/` | CDM 自动化模块（`modAutoImportNest.bas` 一键导入+排版，`Events.bas` 菜单注册） |
| `AI工作流程/` | CDM 源码分析（Make/Events/frmNTCW）+ 自动化方案文档 |
| `软件工作流程.md` | AI 操作手册：CSV 导入 → 批量生产 → 排版的完整工作流 |
| `open_vba_editor.py` | 自动激活/最大化 AlphaCAM 窗口并打开 VBA 编辑器（Alt+F11） |
| `RevNest_source/` | RevNest 反向排版 v1.2 插件完整源码（从 AlphaCAM 提取） |
| `install.bat` | Windows 一键安装脚本 |
| `install_vba.py` | VBA 代码安装到 AlphaCAM 的 Python 脚本 |
| `make_icons.py` | 生成工具栏 BMP 图标的工具 |
| `SKILL.md` | MCP 技能定义 |
| `requirements.txt` | Python 依赖 |

## 安装

### 前提条件

- Windows 7+ / 10 / 11
- AlphaCAM 2016 R1 已安装
- Python 3.10+（需要 `pywin32` 和 `mcp` 库）

### 快速安装

```bash
# 安装 Python 依赖
pip install -r requirements.txt

# 双击运行 install.bat，或手动注册到 AI 客户端的 MCP 配置
```

### 手动配置

在 AI 客户端的 MCP 配置文件中添加：

```json
{
  "mcpServers": {
    "alphacam-bridge": {
      "command": "python",
      "args": [
        "C:\path\to\server.py",
        "--progid",
        "aroutaps.Application"
      ]
    }
  }
}
```

## MCP 工具清单（69 个）

### 状态与信息（2）
| 工具 | 说明 |
|---|---|
| `get_status` | 检查 AlphaCAM 连接状态、版本、路径 |
| `get_drawing_info` | 获取当前图纸详情（几何数、路径数、图层、操作数） |

### VBA 与插件（9）
| 工具 | 说明 |
|---|---|
| `run_vba_macro` | 运行 VBA 宏（支持传参） |
| `run_vba_line` | 直接执行一行 VBA 代码（自动创建临时模块） |
| `list_vba_modules` | 列出 VBA 项目中所有模块名称和类型 |
| `get_vba_code` | 读取指定 VBA 模块的完整源代码 |
| `install_vba_module` | 安装 VBA 模块（从源码，覆盖同名模块） |
| `delete_vba_module` | **删除 VBA 模块**（按名称清理多余模块） |
| `load_addin` | 加载插件 DLL / VBA 项目 |
| `enable_addin` | 启用/禁用插件 |
| `list_addins` | 列出已加载的全部插件 |

### 路径信息与变换（3）
| 工具 | 说明 |
|---|---|
| `get_path_info` | 读取路径详细信息（包围盒、元素数、前10个元素的端点坐标） |
| `move_path` | 移动路径（局部 MoveL 或全局 MoveG） |
| `rotate_path` | 旋转路径（指定角度和中心点） |

### 文件操作（7）
| 工具 | 说明 |
|---|---|
| `new_drawing` | 新建空白图纸 |
| `open_drawing` | 打开 `.amd` 文件 |
| `open_dxf` | 打开 DXF/DWG 文件 |
| `open_step` | 打开 STEP 文件 |
| `open_stl` | 打开 STL 文件 |
| `save_drawing` | 保存图纸（可另存为） |
| `output_nc` | 输出 NC 代码到文件 |

### 几何创建（7）
| 工具 | 说明 |
|---|---|
| `create_rectangle` | 创建矩形（指定两角点） |
| `create_circle` | 创建圆（指定直径和圆心） |
| `create_circle_3pts` | 创建圆（指定三点） |
| `create_line` | 创建直线（指定两端点） |
| `create_polygon` | 创建正多边形（指定边数/半径） |
| `create_ellipse` | 创建椭圆（指定长/短轴） |
| `create_text` | 创建文字标注 |

### 几何查询（1）
| 工具 | 说明 |
|---|---|
| `get_all_geometries` | 列出全部几何的完整信息（类型、范围、属性） |

### 查询列表（3）
| 工具 | 说明 |
|---|---|
| `list_geometries` | 列出所有几何图形 |
| `list_operations` | 列出所有操作 |
| `list_toolpaths` | 列出所有刀具路径 |

### 删除（2）
| 工具 | 说明 |
|---|---|
| `delete_selected` | 删除选中的几何 |
| `delete_all_geometries` | 删除全部几何（保留刀具路径） |

### 裁剪（1）
| 工具 | 说明 |
|---|---|
| `trim_with_boundary` | 以边界裁剪线段 |

### 加工操作（1）
| 工具 | 说明 |
|---|---|
| `run_machining` | 执行加工（粗精加工、轮廓铣槽、雕刻、钻孔、啄钻、攻丝、镗孔），完整控制进给/转速/深度 |

### 刀具（2）
| 工具 | 说明 |
|---|---|
| `select_tool` | 从库中选择刀具 |
| `get_current_tool` | 获取当前刀具信息 |

### 工作平面与图层（3）
| 工具 | 说明 |
|---|---|
| `create_workplane` | 创建工作平面 |
| `set_workplane` | 设置当前工作平面 |
| `create_layer` | 创建或获取图层（可设颜色） |

### 视图控制（5）
| 工具 | 说明 |
|---|---|
| `view_zoom_extents` | 缩放全图适应屏幕 |
| `view_zoom_window` | 框选窗口缩放 |
| `view_set_direction` | 设置 3D 视角方向（俯视/前视/等轴测等） |
| `lock_acam` | 禁用屏幕刷新（批量操作时加速） |
| `unlock_acam` | 恢复屏幕刷新（可选缩放全图） |

### 材料（1）
| 工具 | 说明 |
|---|---|
| `get_material` | 获取当前图纸材料信息（名称、密度、进给、转速等） |

### 路径操作（5）
| 工具 | 说明 |
|---|---|
| `mirror_path` | 沿直线镜像路径（几何/刀具路径） |
| `offset_path` | 偏置封闭路径（左/右侧，指定距离） |
| `copy_temporary_store` | 复制路径为临时几何，可镜像后存储 |
| `get_path_attributes` | 读取路径的用户属性 |
| `set_path_attribute` | 设置路径的用户属性 |

### 排版信息（3）
| 工具 | 说明 |
|---|---|
| `has_nesting` | 检查当前图纸是否包含排版信息 |
| `get_nesting_info` | 获取排版详情（Sheet、零件、实例数据） |
| `get_sheet_extents` | 获取全部排版 Sheet 的全局范围 |

### 操作排序（2）
| 工具 | 说明 |
|---|---|
| `order_operations_all` | 按排版 Sheet 顺序自动排序所有刀具路径 |
| `order_manual` | 手动指定顺序重排几何/刀具路径 |


### 后处理器（1）
| 工具 | 说明 |
|---|---|
| `select_post` | 选择后处理器 |

### API 文档（5）
| 工具 | 说明 |
|---|---|
| `list_docs` | 列出文档来源目录及分类，**含 CHM 索引状态**（显示全部 33 个 .chm 已转换/待转换状态） |
| `search_docs` | 按关键词搜索，**支持 CHM 描述搜索 + 内容片段预览 + 60 秒缓存** |
| `read_doc` | 读取指定文档页全文（支持 max_len 控制 token 消耗） |
| `chm_to_html` | 单文件转换：将 `.chm` 编译帮助文件转换为 HTML |
| `chm_to_html_all` | **批量转换**：一键转换全部未处理的 20 个 .chm 文件到 `chm/{Key}_html/` |

> 文档搜索自动检测 AlphaCAM 安装目录，覆盖 **tempacamapi**（VBA API）、**ACAM3**（3D 模块）、**ACAM4**（4 轴模块）等所有已提取的 HTML 文档。
>
> CHM 索引系统注册了 **33 个唯一 .chm 帮助文件**，包括 CDM（橱柜设计模块）、APM、ModuleWorks 5 轴加工等。搜索时会同时匹配 CHM 描述和已解压的 HTML 内容。
>
> 使用 `chm_to_html_all` 可一键解压全部 .chm，解压后的 `_html` 目录自动纳入文档搜索路径。

### 实用工具（6）
| 工具 | 说明 |
|---|---|
| `set_undo_point` | 设置撤销点 |
| `zoom_all` | 缩放全图 |
| `run_workflow` | 批量执行多步骤工作流 |
| `list_layers` | 列出所有图层（颜色、可见性） |
| `close_drawing` | 关闭当前图纸（不保存） |
| `shell_command` | 执行系统命令或脚本 |

## VBA 插件功能

> **关于 .arb 文件**：AlphaCAM 的插件以 `.arb`（Add-in Resource Bundle）格式发布，它是一个包含 VBA 源码、窗体、图标和菜单定义的资源包。**.arb 文件必须先被 AlphaCAM 加载（通过菜单或 `load_addin` 工具）**，然后才能通过 `list_vba_modules` 列出模块、通过 `get_vba_code` 读取源码。MCP 无法直接解析 .arb 文件格式，必须经由 AlphaCAM 的 VBA 编辑器接口间接读取。

`CCC功能/` 目录包含 VBA 工具（通过 AlphaCAM 菜单栏 "CCC功能" 访问）：

| 文件 | 功能 |
|---|---|
| `Events.bas` | 插件入口，注册"CCC功能"菜单 |
| `modTrim.bas` | **依边界裁剪** — 选择边界和线段，将线段超出边界的部分裁剪 |
| `modOffset.bas` | **全排版刀具偏移** — 按刀具名称选择，整体偏移 X/Y/Z |
| `modSort.bas` | **排版刀具排序** — 按加工方式+刀具分组，拖拽调整加工顺序 |
| `modMirror.bas` | **反面镜像** — 自动镜像排版 Sheet 几何生成反面（X 轴或 Y 轴镜像） |
| `frmToolOffset.txt` / `frmToolSort.txt` | 刀具偏移/排序对话框的窗体定义 |

### CDM 橱柜门自动化

`CDM功能/` 目录包含 CDM（Cabinet Door Manufacturing）自动化模块：

| 文件 | 功能 |
|---|---|
| `modAutoImportNest.bas` | **自动化生产排版**：弹窗选 CSV → 导入门板数据 → `g_Make_Master` 批量生产+排版 |
| `Events.bas` | CDM 工程菜单注册（含 "自动化生产排版" 按钮） |

自动化流程（已在 CDM 工程中运行验证）：

```
选择 CSV 文件（带记忆）
  → 客户名"自动化生产"（自动创建）
  → 创建订单（重名直接取消）
  → 逐行导入门板明细
     ├── 门型已存在 → 用其 UserStyle（900/930）
     │     930 门型（如平板PETA）→ 复制 UserStyleName + UserValue_0~6
     │     正确加载用户样式宏（AD_OnePanelSquare 等）
     └── 新门型 → 创建为 900 标准镶板门
  → g_Make_Master 批量生产 + 排版 + NC 输出
```

> **关键技术点**：930 用户自定义门型的几何由 VBA 宏生成，导入时必须复制
> `StyleName=UserStyleName`（宏项目名）和 `UserValue_0~6`（宏参数），否则报"无法连接用户定义的宏"。

### RevNest 反向排版

`RevNest_source/` 目录包含从 AlphaCAM 2016 R1 `ReverseNest.arb` 插件提取的完整源码，
实现排版零件的反面镜像生成。详见 [`RevNest_API参考.md`](RevNest_API参考.md)。

## 操作规范

### 🪟 窗口管理（最高优先级）

| 规则 | 说明 |
|---|---|
| ✅ **允许** 调整图形视图窗口 | `view_zoom_extents`、`view_zoom_window`、`view_set_direction`、`zoom_all` |
| ❌ **禁止** 改变主窗口大小 | 绝不设置 `Width` / `Height` |
| ❌ **禁止** 移动主窗口位置 | 绝不设置 `Left` / `Top` |
| ❌ **禁止** 改变主窗口状态 | 绝不设置 `WindowState`（最大化/最小化/还原） |
| ⚠️ `Visible` 仅按需设置 | 仅在连接时根据 `ALPHACAM_VISIBLE` 环境变量设置一次 |
| 🛡️ 释放 COM 时不改窗口 | `_cleanup()` 不会设置 `Visible` 或任何窗口属性 |

所有视图缩放/方向操作均通过 `Drawing.ViewWindow` 进行，不影响主窗口布局。

### 🔌 连接规范

| 规则 | 说明 |
|---|---|
| 自动重连 | COM 断开后指数退避重连（最多 5 次，间隔 1s→2s→4s→8s→16s） |
| 僵尸检测 | `_check_alive()` 双层探测（Name + ActiveDrawing），防止误判 |
| 状态回调 | 连接状态变化通过 `set_state_callback()` 通知 |
| 多实例支持 | 通过 `--progid` 参数指定不同 ProgID 切换实例 |
| VBA 工程定位 | `_get_vba_project()` 自动定位：ActiveVBProject → CDM 工程 → 首个工程 |

### 🖥️ VBA 编辑器自动化

用户已授权 AI 在需要时自动打开 VBA 编辑器：

| 操作 | 方式 |
|---|---|
| 激活+最大化窗口 | `open_vba_editor.py`（user32 SetForegroundWindow + ShowWindow SW_MAXIMIZE） |
| 打开 VBA 编辑器 | 发送 Alt+F11 快捷键 |
| 读取/写入模块 | 直接通过 `get_vba_code` / `install_vba_module`，无需打开编辑器 |

### 📐 几何操作规范

| 类型 | 空间 |
|---|---|
| 几何图元 | 在工作平面（Workplane）坐标系中创建 |
| 图形视图 | 可通过 `view_*` 工具缩放/平移/旋转视角，不影响实际坐标 |
| 图层 | 通过 `create_layer` 创建和命名图层，支持 RGB 颜色设置 |

- GitHub: https://github.com/cczzyy-cn/AlphacamMCP
- 问题反馈: https://github.com/cczzyy-cn/AlphacamMCP/issues
