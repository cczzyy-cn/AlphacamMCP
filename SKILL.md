---
name: alphacam-bridge
description: 通过 MCP 协议直接实时操作 AlphaCAM 2016 R1 的桥接器插件使用指南。
---

# AlphaCAM MCP 桥接器 — AI 实时操作 AlphaCAM

通过 `alphacam-bridge` 插件提供 **69 个 MCP 工具**，用于直接控制 AlphaCAM 2016 R1。

## 可用工具

### 状态与信息
- **`get_status`** — 检查连接状态、版本、路径
- **`get_drawing_info`** — 当前图纸详情（几何、路径、图层、操作）

### 文件
- **`new_drawing`** / **`open_drawing`** / **`open_dxf`** / **`open_step`** / **`open_stl`**
- **`save_drawing`** / **`output_nc`**

### 几何创建与查询
- **`create_rectangle`** / **`create_circle`** / **`create_line`** / **`create_polygon`** / **`create_ellipse`** / **`create_text`**
- **`get_all_geometries`** / **`list_geometries`** / **`list_operations`** / **`list_toolpaths`**
- **`delete_selected`** / **`delete_all_geometries`**

### 裁剪与加工
- **`trim_with_boundary`** — 以边界裁剪线段
- **`run_machining`** — 完整加工控制（进给、转速、深度、加工类型）

### 工作平面与图层
- **`create_workplane`** / **`set_workplane`** / **`create_layer`**

### 刀具
- **`select_tool`** / **`get_current_tool`**

### 路径操作
- **`mirror_path`** / **`offset_path`** / **`copy_temporary_store`**
- **`get_path_attributes`** / **`set_path_attribute`**

### 排版
- **`has_nesting`** / **`get_nesting_info`** / **`get_sheet_extents`**
- **`order_operations_all`** / **`order_manual`**

### 屏幕控制
- **`lock_acam`** / **`unlock_acam`**

### VBA 与插件（9）
- **`run_vba_macro`** — 运行 VBA 宏（支持传参）
- **`run_vba_line`** — 直接执行一行 VBA 代码（自动创建临时模块）
- **`list_vba_modules`** — 列出 VBA 项目全部模块
- **`get_vba_code`** — 读取 VBA 模块完整源码
- **`install_vba_module`** — 安装/覆盖 VBA 模块（从源码）
- **`delete_vba_module`** — 删除 VBA 模块（按名称）
- **`load_addin`** / **`enable_addin`** / **`list_addins`**

### 后处理器
- **`select_post`**

### API 文档
- **`list_docs`** — 列出文档来源 + CHM 索引状态（33 个 .chm 转换状态）
- **`search_docs`** — 关键词搜索（含 CHM 描述搜索 + 内容片段预览 + 60s 缓存）
- **`read_doc`** — 读取文档页全文（max_len 控制）
- **`chm_to_html`** — 单文件 CHM 转 HTML
- **`chm_to_html_all`** — 批量转换全部未处理 .chm

### 实用工具
- **`set_undo_point`** / **`zoom_all`** / **`run_workflow`**（批量多步骤）
- **`list_layers`** / **`close_drawing`** / **`shell_command`**

## 加工类型常量
- 1 = 粗精加工
- 2 = 轮廓铣槽
- 10 = 雕刻
- 21 = 钻孔 / 22 = 啄钻 / 23 = 攻丝 / 24 = 镗孔

## CDM 橱柜门自动化（modAutoImportNest）

CDM 工程中已安装 `modAutoImportNest` 模块，实现一键自动化：

```
选择 CSV 文件（带记忆）→ 客户"自动化生产" → 创建订单（重名取消）
  → 逐行导入门板明细（复制 UserStyle 参数）
  → g_Make_Master 批量生产 + 排版 + NC 输出
```

- 菜单触发：AlphaCAM 菜单 → CDM → 自动化生产排版（宏 `m_AutoImportNest`）
- 930 用户样式门型（如平板PETA）需复制 `StyleName=UserStyleName` + `UserValue_0~6`，否则报"无法连接用户定义的宏"
- VBA 工程自动定位：ActiveVBProject → CDM 工程 → 首个工程（`_get_vba_project()`）

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
| 安装/读取/删除 VBA 模块 | `install_vba_module` / `get_vba_code` / `delete_vba_module` |

## 注意事项

- 桥接器通过 COM 连接 AlphaCAM，AlphaCAM 必须在运行状态
- VBA 宏调用使用 `Project.Module.Procedure` 格式
- 工作平面检查：打开临时图纸时检查 `Drawing.WorkPlanes.count`，若包含工作平面则无法插入
- 文档搜索自动检测 AlphaCAM 安装目录，涵盖 VBA API + 3D/4D 用户手册
- 支持 `.env` 文件配置 `ALPHACAM_PROG_ID`、`ALPHACAM_VISIBLE` 等环境变量
- 需要打开 VBA 编辑器时：运行 `open_vba_editor.py`（激活/最大化 AlphaCAM 窗口 + Alt+F11），已获用户授权
