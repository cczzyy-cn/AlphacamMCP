# CDM功能 — CDM 橱柜门自动化

AlphaCAM CDM（Cabinet Door Manufacturing）自动化模块源码与文档。

## 文件说明

| 文件 | 说明 |
|------|------|
| `modAutoImportNest.bas` | ⭐ **自动化生产排版模块**：导入门板数据 → `g_Make_Master` 批量生产+排版+NC 输出；含菜单入口 `AutoImportNest`（弹 `frmAutoNest` 窗体）与带参入口 `AutoImportNestWithParams` |
| `frmAutoNest.txt` | ⭐ **自动化生产排版窗体**（代码文本）：CSV 路径记忆回填 + 系统文件对话框 + 确定/取消（AlphaCAM 不支持导入 .frm，需手动创建，见 `frmAutoNest_手动创建.md`） |
| `Events.bas` | CDM 工程菜单注册（含"自动化生产排版"按钮 + `m_AutoImportNest` 包装函数） |
| `Make.bas` | CDM 原始 Make 模块源码备份（7926 行，加工引擎） |
| `CDM分析报告.md` | CDM 完整源码分析（模块结构、数据库表、调用链） |
| `README.md` | 本文档 |

## 自动化生产排版（modAutoImportNest）

### 使用方式

```
AlphaCAM 菜单 → CCC功能 → 自动化生产排版   （弹出 frmAutoNest 窗体）
```

- 菜单项绑定 `Events.bas` 的 `m_AutoImportNest` → `modAutoImportNest.AutoImportNest`（弹出 `frmAutoNest` 窗体）
- 窗体"确定"→ `AutoImportNestWithParams(路径, "自动化生产", 材料, bRunNest)`：
  - 不勾选"只导入订单，不生产排版" → `bRunNest=True`（导入 + 排版）
  - 勾选 → `bRunNest=False`（仅导入订单，跳过排版）

### 完整流程

```
1. 弹出 frmAutoNest 窗体：CSV 路径记忆回填（注册表 CCC\AutoImportNest\LastPath），
   可点"..."用系统文件对话框选择，或直接输入路径
2. 确定 → 客户名"自动化生产"（不存在自动创建）
3. 创建订单（订单名已存在 → 直接取消并提示）
4. 逐行导入门板明细：
   ├── 门型已存在 → 用其 UserStyle 判断 StyleNumber
   │     ├── UserStyle=True  → 930（用户自定义门型）
   │     │    复制 UserStyleName + UserVariableString + UserValue_0~6
   │     │    （否则宏调用失败："无法连接用户定义的宏"）
   │     └── UserStyle=False → 900（标准镶板门）
   └── 新门型 → 自动创建为 900 标准镶板门
5. 材料不存在 → 自动创建（默认"开料机3000mm" 18×1220×3000）
6. 调用 g_Make_Master(OrderID) → 批量生产 + 排版 + NC
```

### CSV 字段映射（1-based 列号）

| CSV 列 | 0-based | 数据库字段 | 说明 |
|--------|:-------:|-----------|------|
| 列1 造型名称 | 0 | `TypeName` / `StyleName`(回退) | 门板类型 |
| 列2 宽 | 1 | `Width` | |
| 列3 高 | 2 | `Length` | |
| 列4 数量 | 3 | `Quantity` | |
| 列5 颜色 | 4 | `CSV_ItemNumber` + `ComponentGrouping`(Val) | 颜色文本；组编号转数字 |
| 列6 客户名 | 5 | `CSV_CustomerName` | |
| 列8 开启方向 | 7 | `CustomField1` | |
| 列9 终端地址 | 8 | `CustomField2` | |
| 列10 板件码 | 9 | `CSV_OrderNumber` | 订单号 |
| 列12 备注 | 11 | `ProductionComment` | |
| 列13 材料 | 12 | `Material` | 空时用默认"开料机3000mm" |

### 关键技术点

| 要点 | 说明 |
|------|------|
| **930 门型宏参数** | `AD_ORDER_DETAILS.StyleName` 必须 = `AD_DOOR_TYPES.UserStyleName`（宏项目名，如 `AD_OnePanelSquare`），否则 `gbln_ProjectExists` 找不到宏报错 |
| **UserValue_0~6** | INSERT...SELECT 从 `AD_DOOR_TYPES` 直接复制，供 `App.Run` 传参给宏 |
| **ComponentGrouping 类型** | Long 整数，CSV 颜色文本需 `Val()` 转换（文本→0） |
| **订单重名** | 已存在则直接取消导入（不弹窗询问） |
| **材料默认** | "开料机3000mm"：18mm 厚、1220×3000mm 板 |

### 安装方式

```python
# 通过 MCP 安装到 CDM 工程
code = open("CDM功能/modAutoImportNest.bas", encoding="gbk").read()
install_vba_module(module_name="modAutoImportNest", code=code)

# Events.bas 已含菜单注册（m_AutoImportNest 包装函数），若菜单缺失需覆盖导入
code = open("CDM功能/Events.bas", encoding="gbk").read()
install_vba_module(module_name="Events", code=code)
```

> **frmAutoNest 窗体**：AlphaCAM VBA 不支持导入 .frm 设计文件，需按
> `frmAutoNest_手动创建.md` 手动创建窗体与 6 个控件（含"只导入订单"勾选框），
> 再粘贴 `frmAutoNest.txt` 代码（窗体代码若更新，在 VBA 编辑器中整体替换代码窗口内容即可）。

> 若 `install_vba_module` 报"工程已被保护"，需先在 VBA 编辑器确认工程未锁定，
> 或先 `delete_vba_module` 删除同名旧模块再安装。

## 数据库结构速查

核心表：`AD_ORDERS`（订单）、`AD_ORDER_DETAILS`（门板明细）、`AD_DOOR_TYPES`（门型）、`AD_DOOR_PATHS`（刀路）、`AD_MATERIALS`（材料）。

详见 `CDM分析报告.md` 第三节。
