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
| `server.py` | MCP 桥接器主程序（Python），通过 STDIO 协议与 AI 通信 |
| `alphacam_com.py` | AlphaCAM COM 自动化封装层 |
| `CCC功能合集.bas` | VBA 插件模块（依边界裁剪、全排版刀具偏移、排版刀具排序） |
| `frmToolOffset.txt` | 刀具偏移对话框的窗体定义（VBA UserForm） |
| `frmToolSort.txt` | 刀具排序对话框的窗体定义（VBA UserForm） |
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

## MCP 工具清单（48 个）

### 状态与信息（2）
| 工具 | 说明 |
|---|---|
| `get_status` | 检查 AlphaCAM 连接状态、版本、路径 |
| `get_drawing_info` | 获取当前图纸详情（几何数、路径数、图层、操作数） |

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

### 几何创建（6）
| 工具 | 说明 |
|---|---|
| `create_rectangle` | 创建矩形（指定两角点） |
| `create_circle` | 创建圆（指定直径和圆心） |
| `create_line` | 创建直线（指定两端点） |
| `create_polygon` | 创建正多边形（指定边数/半径） |
| `create_ellipse` | 创建椭圆（指定长/短轴） |
| `create_text` | 创建文字标注 |

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

### VBA 与插件（3）
| 工具 | 说明 |
|---|---|
| `run_vba_macro` | 运行 VBA 宏（支持传参） |
| `load_addin` | 加载插件 DLL / VBA 项目 |
| `enable_addin` | 启用/禁用插件 |

### 后处理器（1）
| 工具 | 说明 |
|---|---|
| `select_post` | 选择后处理器 |

### API 文档（3）
| 工具 | 说明 |
|---|---|
| `list_docs` | 列出 API 文档分类及文件数 |
| `read_doc` | 读取指定 API 文档页 |
| `search_docs` | 搜索 API 文档 |

### 实用工具（3）
| 工具 | 说明 |
|---|---|
| `set_undo_point` | 设置撤销点 |
| `zoom_all` | 缩放全图 |
| `run_workflow` | 批量执行多步骤工作流 |

---

## VBA 插件功能

## VBA 插件功能

`CCC功能合集.bas` 包含三个工具（通过 AlphaCAM 菜单栏 "CCC功能" 访问）：

1. **依边界裁剪** — 选择边界和线段，将线段超出边界的部分裁剪
2. **全排版刀具偏移** — 按刀具名称选择，整体偏移 X/Y/Z
3. **排版刀具排序** — 按加工方式+刀具分组，拖拽调整加工顺序

## 项目链接

- GitHub: https://github.com/cczzyy-cn/AlphacamMCP
- 问题反馈: https://github.com/cczzyy-cn/AlphacamMCP/issues
