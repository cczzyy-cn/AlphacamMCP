# CDM功能 — CDM 橱柜门自动化

AlphaCAM CDM（Cabinet Door Manufacturing）自动化模块源码与文档。

## 文件说明

| 文件 | 说明 |
|------|------|
| `modAutoImportNest.bas` | ⭐ **自动化生产排版模块**（v1.10，1033 行）：导入门板数据 → `g_Make_Master` 批量生产+排版+NC 输出；含菜单入口 `AutoImportNest`（弹 `frmAutoNest` 窗体）与带参入口 `AutoImportNestWithParams`（v1.10 起可传窗体选中的材料，整批覆盖 CSV 材料列）；**门板标签 EMF 重生成**（稳定唯一码对齐报表行） |
| `frmAutoNest.txt` | ⭐ **自动化生产排版窗体**（代码文本）：CSV 路径记忆回填 + 系统文件对话框 + **材料下拉**（v1.10，启动时从 `AD_MATERIALS` 加载）+ 确定/取消（AlphaCAM 不支持导入 .frm，需手动创建，见 `frmAutoNest_手动创建.md`） |
| `Events.bas` | CDM 工程菜单注册（`Events.bas:277` 注册"自动化生产排版"按钮 → `m_AutoImportNest` 包装函数；`mint_UpdateDB` 里含 `AD_REPORT_DATA.PressPieceUID` 建列） |
| `Make.bas` | CDM 原始 Make 模块源码（8005 行，加工引擎；v2.2 起在打件序号时同步写稳定唯一码 `DEF_ATT_PIECE_UID`） |
| `modAutoImportNest分析.md` | 该模块的完整分析（5.3 标签行身份、重复标签根因链与 A/B 修法） |
| `CDM数据库说明.md` | `CDM.mdb` 表结构与字段说明（含 `AD_REPORT_DATA` 行身份、DDL 锁限制） |
| `Make.bas分析.md` | `Make.bas` 的模块结构与调用链分析 |
| `CDM分析报告.md` | CDM 完整源码分析（模块结构、数据库表、调用链） |
| `标签重生成优化方案.md` | 标签重生成的设计与优化记录 |
| `frmAutoNest_手动创建.md` | 窗体手动创建步骤（AlphaCAM 不支持 .frm 导入） |
| `README.md` | 本文档 |

> **改这三个 `.bas` 请走仓库根的 `tools/` 部署闭环**（快照 → 三方审计 → 部署读回校验 →
> 编译探针），见 [`../tools/README.md`](../tools/README.md)。改动前先 `running_snapshot.py` 存基线。

## 自动化生产排版（modAutoImportNest）

### 使用方式

```
AlphaCAM 菜单 → CDM → 自动化生产排版   （弹出 frmAutoNest 窗体）
```

- 菜单项绑定 `Events.bas` 的 `m_AutoImportNest` → `modAutoImportNest.AutoImportNest`（弹出 `frmAutoNest` 窗体）
- 窗体"确定"→ `AutoImportNestWithParams(CSV路径, "自动化生产", bRunNest, bOverwrite, sMaterialOverride)`：
  - **材料来自窗体下拉**（v1.10）：`sMaterialOverride` 非空 → 整批统一用该材料，**忽略 CSV 第 13 列**；为空 → 逐行取自 CSV 第 13 列（0 基 12）。两条路径都要求材料已存在于 `AD_MATERIALS`
  - 不勾选"只导入订单，不生产排版" → `bRunNest=True`（导入 + 排版）
  - 勾选 → `bRunNest=False`（仅导入订单，跳过排版）
  - 勾选"强制覆盖重名订单" → `bOverwrite=True`：订单名已存在时删除原订单相关数据（`AD_ORDER_DETAILS`、`AD_REPORT_DATA`、`AD_ORDERS`）后重新导入

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
5. 材料校验：窗体选中的材料（或逐行的 CSV 材料）必须已存在于 AD_MATERIALS 表，缺失则整体失败并回滚
   （v1.6 起只校验、不自动建档，避免静默建出错规格材料）
6. 调用 g_Make_Master(OrderID) → 批量生产 + 排版 + NC
```

> **v1.7 事务与顺序**：CSV 文件存在性检查提前到「创建订单」之前（路径写错不再
> 留下空订单 + 客户记录）；客户/订单/门型/明细全部纳入**同一个数据库事务**
> （`BeginTrans`/`CommitTrans`/`RollbackTrans`），任一步失败整体回滚 ——
> 覆盖模式下「先删旧订单、再插新明细」中途失败时，旧数据随回滚恢复，不会丢。

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
| 列13 材料 | 12 | `Material` | **必填**：必须已存在于 `AD_MATERIALS`，否则整单导入失败回滚；**v1.10 起窗体选中的材料会整批覆盖本列** |

### 关键技术点

| 要点 | 说明 |
|------|------|
| **930 门型宏参数** | `AD_ORDER_DETAILS.StyleName` 必须 = `AD_DOOR_TYPES.UserStyleName`（宏项目名，如 `AD_OnePanelSquare`），否则 `gbln_ProjectExists` 找不到宏报错 |
| **UserValue_0~6** | INSERT...SELECT 从 `AD_DOOR_TYPES` 直接复制，供 `App.Run` 传参给宏 |
| **ComponentGrouping 类型** | Long 整数，CSV 颜色文本需 `Val()` 转换（文本→0） |
| **订单重名** | 默认直接取消导入（不弹窗询问）；勾选窗体"强制覆盖重名订单"则删除原订单明细/报表后重建 |
| **材料校验** | 材料按名查 `AD_MATERIALS`，不存在即报错并回滚整单；不自动建档（v1.6 起改为只校验，v1.7 删除遗留的自动建档死代码 `glng_EnsureMaterial`）。v1.10 起窗体材料在事务开始前**一次性校验**，再整批覆盖明细材料 |

### 重新生成门板标签（g_RegenDoorLabelEMFs）

排版后若在 ARD 嵌套图里**手动移动过门板**，用窗体上的"重新生成标签"按钮重出标签图：

- 入口：`frmAutoNest.cmdRegenLabel_Click` → `modAutoImportNest.g_RegenDoorLabelEMFs`
- 前提：当前图纸是排版后的嵌套档案（`GetNestInformation` 非空），且刀路含 `DEF_ATT_JOB_NAME` 属性
- 材料名四级回退：图纸名 `<Job>_<材料>.ard` → `AD_REPORT_DATA.PressDoorImage` 路径 →
  `AD_ORDER_DETAILS.Material`（多材料则放弃）→ 图纸属性 `DEF_ATT_MATERIAL_NAME`
  （嵌套板的 `MaterialName` 是板名配置如 "Admin"，**不是材料名**，不能直接取）
- 产出：`<Job>_<材料>_<板名>_<件号>.emf`，并同步 `AD_REPORT_DATA.PressDoorImage` / `PressDoorCounter`
- 保护原图：把当前图 `SaveAs` 到临时副本后在副本上生成，用户正式图纸数据零触碰（保住加工道次关联）

> **v1.7 健壮性修复**：
> ① 开始前弹「未保存修改」确认 —— 未保存的手动移动只用于出标签、不会写回原图，重生成后会丢失；
> ② 失败时自动 *回滚数据库 + 从备份目录还原旧 EMF + 清理 `regen_*` 临时文件 + 重新打开用户原档案*
> （原实现失败后会把用户留在空白图/临时副本里）；
> ③ 等待 EMF 落盘的期望值改为「板数 + 总件数」（原实现只传总件数，而目录通配计数会把
> 整板图 `<Job>_<材料>_<板名>.emf` 一并计入，可能提前返回）。

> **v1.8 修复（2026-09-10 实机定位）**：
> ① **屏幕刷新泄漏**（"窗口标题老显示 `regen_<订单>_<Timer>`"的根因）——
> `Make.m_CreateAlphaCAMDrawingsOfSheets` 会置 `ScreenUpdating=False` /
> `ProjectBarUpdating=False`，而收尾处的恢复语句在 `Make.bas:3991/3992` **被注释掉了**，
> 于是每次生产/重生成后 AlphaCAM 不再重绘，标题与画面停在临时副本名上
> （此时 `ActiveDrawing.FullName` 其实已是真档案、`Modified=False`）。已恢复那两行，
> 并在本模块成功/失败路径都兜底 `ScreenUpdating=True` + `Redraw`。
> ② **备份目录改为先 `Kill` 再 `RmDir`**：旧 EMF 让目录非空，`RmDir` 必定失败且被静默吞掉，
> 曾累积 28 个 `regen_backup_*`（222 个旧 EMF）。
> ③ **临时嵌套 ard 改在 `App.New` 关档后补删并记日志**（原先删除时文件仍被占用，静默失败）。

> **v1.10（2026-09-13）窗体材料下拉**：`frmAutoNest` 新增 `lblMaterial` + `cboMaterial`，
> 启动时从 `AD_MATERIALS` 读全部材料（`Style=2` 只可选，防手输错名），预选顺序 =
> 上次选择（注册表 `CCC\AutoImportNest\LastMaterial`）→ `MaterialDefault=True` 的行 → 第一项。
> 选中的材料**对整批所有行生效并忽略 CSV 第 13 列**；模块侧新增可选形参
> `sMaterialOverride`（空 = 保留旧的逐行 CSV 行为），在事务开始前一次性校验材料存在性，
> 校验不过直接失败、不产生任何数据库改动。
>
> 窗体控件是通过 VBIDE `Designer.Controls.Add` 自动加的（**必须先加控件、再部署引用它的代码**，
> 反了会编译报"找不到方法或数据成员"）；材料行插在 CSV 行下方，原 Top ≥ 66 的控件整体下移 24。
> 改动后记得让 AlphaCAM 正常保存/退出，把 VBA 工程写回 `CDM.arb`。

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
> `frmAutoNest_手动创建.md` 手动创建窗体与 10 个控件（含 v1.10 的材料下拉 `lblMaterial`/`cboMaterial`、"只导入订单"/"强制覆盖重名订单"勾选框与"重新生成标签"按钮），
> 再粘贴 `frmAutoNest.txt` 代码（窗体代码若更新，在 VBA 编辑器中整体替换代码窗口内容即可）。

> 若 `install_vba_module` 报"工程已被保护"，需先在 VBA 编辑器确认工程未锁定，
> 或先 `delete_vba_module` 删除同名旧模块再安装。

## 数据库结构速查

核心表：`AD_ORDERS`（订单）、`AD_ORDER_DETAILS`（门板明细）、`AD_DOOR_TYPES`（门型）、`AD_DOOR_PATHS`（刀路）、`AD_MATERIALS`（材料）。

详见 `CDM分析报告.md` 第三节。
