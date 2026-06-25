---
name: alphacam-bridge
description: 通过 MCP 协议直接实时操作 AlphaCAM 2016 R1 的桥接器插件使用指南。
---

# AlphaCAM MCP 桥接器 — AI 实时操作 AlphaCAM

通过 `alphacam-bridge` 插件提供 **49 个 MCP 工具**，用于直接控制 AlphaCAM 2016 R1。

## 可用工具

### 状态与信息
- **`get_status`** — 检查连接状态、版本、路径
- **`get_drawing_info`** — 当前图纸详情（几何、路径、图层、操作）
- **`list_geometries`** / **`list_operations`** / **`list_toolpaths`** — 列出实体

### 文件
- **`new_drawing`** / **`open_drawing`** / **`open_dxf`** / **`open_step`** / **`open_stl`**
- **`save_drawing`**

### 几何创建
- **`create_rectangle`** / **`create_circle`** / **`create_line`** / **`create_polygon`** / **`create_ellipse`** / **`create_text`**

### 工作平面与图层
- **`create_workplane`** / **`set_workplane`** / **`create_layer`**

### 刀具
- **`select_tool`** / **`get_current_tool`**

### 加工
- **`run_machining`** — 完整加工控制（进给、转速、深度、加工类型）
- **`output_nc`** — 输出 NC 代码到文件

### VBA 与插件
- **`run_vba_macro`** — 调用任意 VBA 宏
- **`load_addin`** / **`enable_addin`**

### 实用工具
- **`set_undo_point`** / **`zoom_all`** / **`select_post`**

### 批量操作
- **`run_workflow`** — 链式执行多个步骤

## 加工类型常量
- 1 = 粗精加工
- 2 = 轮廓铣槽
- 10 = 雕刻
- 21 = 钻孔 / 22 = 啄钻 / 23 = 攻丝 / 24 = 镗孔

## MCP 工具 vs VBA 代码 选择指南

| 场景 | 使用 |
|------|------|
| 快速创建几何 | MCP 工具 |
| 标准加工操作 | MCP 工具 |
| 复杂批量操作 | `run_workflow` |
| 自定义逻辑/循环/条件 | 生成 VBA 代码 |
| 自定义菜单/插件 | 生成 VBA 代码 |
| 事件处理（保存前、输出后） | 生成 VBA 代码 |
| 交互式用户选择 | MCP 工具 |
| 多步骤自动化工作流 | MCP `run_workflow` |

## 注意事项

- 桥接器通过 COM 连接 AlphaCAM，AlphaCAM 必须在运行状态
- VBA 宏调用使用 `Project.Module.Procedure` 格式
- 工作平面检查：打开临时图纸时检查 `Drawing.WorkPlanes.count`，若包含工作平面则无法插入
