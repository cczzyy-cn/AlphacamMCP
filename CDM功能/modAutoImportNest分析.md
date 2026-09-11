# modAutoImportNest.bas 分析 —— CDM 自动化生产排版

> 文件: `CDM功能/modAutoImportNest.bas`（**v1.8，2026-09-10；38,877 字节 / 864 行 / GBK 编码 / LF**）
> 定位: CDM（橱柜门制造）工程内的核心 VBA 模块，由「CCC功能」菜单触发，
> 实现 **CSV 订单导入 → 批量生产 → 排版 + NC 输出**，以及排版完成后的
> **门板标签 EMF 重生成**（含失败现场还原）。
>
> 关联窗体: `CDM功能/frmAutoNest.txt`（手动创建，见 `frmAutoNest_手动创建.md`）
> 生产引擎: `CDM功能/Make.bas` 的 `g_Make_Master(OrderID)`

---

## 一、总体结构

模块导出 **3 个 Public 入口** + **11 个 Private 辅助过程**（共 14 个过程）：

| # | 过程 | 类型 | 作用 |
|---|---|---|---|
| 1 | `AutoImportNest()` | Public Sub | 菜单入口，仅以**非模态**方式弹出 `frmAutoNest` |
| 2 | `AutoImportNestWithParams(...)` | Public Sub | **带参主流程入口**，窗体"确定"按钮调用 |
| 3 | `g_RegenDoorLabelEMFs()` | Public Sub | 重生成门板标签 EMF + 事务同步数据库（v1.7 含失败还原） |
| 4 | `ImportCSV` | Private Fn | CSV 解析 + 订单/明细落库（**v1.7 全流程事务**） |
| 5 | `glng_EnsureCustomer` | Private Fn | 查/建客户，返回 CustomerID |
| 6 | `glng_CreateOrder` | Private Fn | 查/建订单；重名按 `bOverwrite` 覆盖或取消 |
| 7 | `glng_EnsureStyle` | Private Fn | 查/建门型，返回 900 / 930 并带回用户样式名 |
| 8 | `m_CheckMaterialExists` | Private Fn | 校验材料存在于 `AD_MATERIALS`（**只校验、不建档**） |
| 9 | `SplitCSVLine` | Private Fn | 手写 CSV 解析器（支持引号与 `""` 转义） |
| 10 | `GetF` | Private Fn | 安全取列（越界返回默认值） |
| 11 | `m_WaitForLabelEMFs` | Private Fn | 轮询等待标签 EMF 落盘（最多 ~6s） |
| 12 | `m_RestoreLabelBackup` | Private Sub | 失败时从备份目录还原旧 EMF |
| 13 | `m_Log` | Private Sub | 日志落盘 `CDM_Import.log` |
| 14 | `m_LogError` | Private Sub | 带步骤名与 `Err.Number` 的错误日志 |

> `glng_EnsureMaterial`（自动建材料）已在 **v1.7 作为死代码删除**。

---

## 二、主流程：AutoImportNestWithParams

```vba
AutoImportNestWithParams(sCSVPath, sCustomerName, bRunNest, Optional bOverwrite = False)
```

> **v1.7 变更**：删除了第 3 个 `sMaterialName` 形参 —— 它只被用于填充一个从未被
> 使用的默认值（死代码）。材料逐行取自 CSV 第 13 列（0 基 12），必须已存在于
> `AD_MATERIALS`。

```
接收 (CSV路径, 客户名, bRunNest, bOverwrite)
    │
    ├─ 从 CSV 路径解析 JobName（取文件名去扩展名，兼容 \ 与 /）
    ├─ gbln_ConnectToDB()                      失败 → 弹错退出
    ├─ lngOrderID = ImportCSV(...)             ┐ -1 = 取消（订单名冲突且不覆盖）
    │                                          ┘  0 = 失败（ImportCSV 内已提示+日志）
    ├─ 若 bRunNest：                           （勾选"只导入订单"则为 False）
    │     Frame.ShowProgressBox「批量生产+排版」
    │     清除注册表残留键 Nest Completed      ← 避免上一订单残留导致误判
    │     Call g_Make_Master(CStr(OrderID))    ← CDM 重头，批量生产 + 排版 + NC
    │     Frame.CloseProgressBox
    │     读回 "Nest Completed" 注册表键判断成功/警告
    └─ 否则仅提示「CSV 导入完成」
```

**关键**：真正的排版引擎是外部 `g_Make_Master`（定义于 `Make.bas:2345`），
本模块只负责**构造数据（订单 + 明细）并调用**。`bRunNest` 布尔开关让用户能
「只导入订单、不生产排版」。

失败判定依赖 `LICOM AlphaDOOR\Nest Parameters\Nest Completed` 注册表键，
属启发式：`g_Make_Master` 中途返回会得到「可能未完全成功」的警告而非错误。

---

## 三、ImportCSV —— 数据导入核心

### 1. 执行顺序（v1.7）

```
① 文件存在性检查            ← v1.7 上移：原先在建订单之后，路径写错会留下空订单+客户
② gdb_CDM.BeginTrans        ← v1.7 新增：整个导入纳入一个事务
③ glng_EnsureCustomer       查 AD_CUSTOMERS，无则 INSERT
④ glng_CreateOrder          查/删/建 AD_ORDERS（重名策略见下）
⑤ 逐行读 CSV → 校验 → glng_EnsureStyle → m_CheckMaterialExists → INSERT 明细
⑥ gdb_CDM.CommitTrans       任一步失败 → EH 回滚 → CleanUp 兜底回滚
```

**事务语义**：客户 / 订单 / 门型（新门型 INSERT）/ 明细全部在同一事务内。
对**覆盖模式**尤其关键 —— 旧订单的 `DELETE` 也在事务里，中途失败时旧数据随回滚恢复，
不会出现「旧的删了、新的没插全」。

### 2. CSV 字段映射（0 基，`GetF(vF, i)`）

| 列 | 变量 | 含义 | 处理 |
|---|---|---|---|
| 0 | `sTp` | **门型** TypeID | 经 `gs_FixSQL` |
| 1 | `w` | 宽 | `Val()`，`≤0` 跳过该行 |
| 2 | `h` | 高 | `Val()`，`≤0` 跳过该行 |
| 3 | `q` | 数量 | `Val()`，`≤0` 归 1 |
| 4 | `sGrp` | 组件分组 | 文本进 `CSV_ItemNumber`，`Val(sGrp)` 进 `ComponentGrouping` |
| 5 | `sCu` | 客户名 | → `CSV_CustomerName` |
| 6 | — | **未使用**（跳过） | |
| 7 | `sC1` | CustomField1 | |
| 8 | `sC2` | CustomField2 | |
| 9 | `sRf` | 柜体编号 | → `CSV_OrderNumber` |
| 11 | `sRm` | 备注 | → `ProductionComment` |
| 12 | `sMat` | **材料** | **必填**，必须存在于 `AD_MATERIALS` |

> 列数 `< 4` 的行跳过；首行按表头丢弃。

### 3. 每行处理

1. **保证门型**：`glng_EnsureStyle(sTp, sUsrStyle)` 返回 StyleNumber 并带回用户样式名：
   - 门型不存在 → `INSERT` 为新门型（UserStyle=False）→ 返回 **900**（标准镶板门）
   - 已存在且 `UserStyle=True` → 返回 **930**（用户自定义样式），带回 `sUsrStyle=UserStyleName`
   - 已存在且非用户样式 → 返回 900
   - `sTp` 为空 → 直接返回 900
2. **校验材料**：`m_CheckMaterialExists(sMat)` 查 `AD_MATERIALS`；
   **不存在或为空 → 先回滚事务再弹错、整单失败**（不做自动建档）。
3. **插入明细**：核心是 **`INSERT ... SELECT` 从 `AD_DOOR_TYPES dt` 复制一行**：

```sql
INSERT INTO AD_ORDER_DETAILS
  (OrderID,TypeName,StyleName,StyleNumber,Quantity,Width,Length,
   Material,ProductionComment,CSV_CustomerName,CSV_OrderNumber,CSV_ItemNumber,
   CustomField1,CustomField2,ComponentGrouping,CornerRadius,RotationMethod,RotationAngle,
   IgnoreOuterGeometry,ByPassNest,UserVariableString,UserDescriptionString,
   UserValue_0..UserValue_6)
SELECT  <OrderID>,'<门型>','<sUsrStyle>',<lngStyleNum>,<q>,<w>,<h>,'<材料>',
        '<备注>','<客户>','<柜号>','<分组>','<C1>','<C2>', Val(sGrp),
        dt.CornerRadius,dt.RotationMethod,dt.RotationAngle,
        dt.IgnoreOuterGeometry,dt.ByPassNest,
        dt.UserVariableString,dt.UserDescriptionString,
        dt.UserValue_0..dt.UserValue_6
FROM AD_DOOR_TYPES dt WHERE dt.TypeID='<门型>'
```

**这是 930 用户自定义样式门型的关键**：`StyleName=UserStyleName`（宏项目名）+ `UserValue_0~6`
（宏参数）必须一起复制，否则冷启动时 AlphaCAM 报 **「无法连接用户定义的宏」**。

### 4. 失败语义

| 情况 | 结果 |
|---|---|
| `INSERT` 受影响行数 = 0（门型在 `AD_DOOR_TYPES` 无参数行） | 计入 `lngFail` 并**继续**（容忍部分失败），循环后汇总提示前 10 条 |
| 材料不存在/为空 | **整单失败 + 回滚** |
| SQL / 运行时错误 | EH：先回滚 → 写日志（含 SQL 前 400 字）→ 提示 → 返回 `ORDER_FAIL` |
| 订单重名且不覆盖 | `glng_CreateOrder` 返回 `ORDER_CANCEL`，回滚（含已建的客户） |

返回码常量：`ORDER_CANCEL = -1`、`ORDER_FAIL = 0`，成功返回 `OrderID`。

---

## 四、辅助函数

### glng_EnsureCustomer / glng_CreateOrder / m_CheckMaterialExists
- 均为「查询 → 无则 INSERT → 返回 ID」模式，新增记录统一用 `@@IDENTITY` 取回主键。
- `glng_CreateOrder`：重名且 `bOverwrite=False` → **返回 -1 取消**并提示；
  `bOverwrite=True` → 级联删除 `AD_ORDER_DETAILS` / `AD_REPORT_DATA` / `AD_ORDERS`
  后重建（删除失败不再静默吞错，交由外层事务回滚）。

### glng_EnsureStyle
- 决定 900 标准 vs 930 用户样式的唯一判断点（见三.3.1）。

### SplitCSVLine —— 手写 CSV 解析器
- 支持双引号包裹字段与 `""` 转义（引号内逗号不拆分）。
- 动态 `ReDim Preserve` 扩容，初始 21 列。

### GetF —— 安全取列
- 越界列返回默认值 `d`，避免下标越界。
- 这是「CSV 列不足 13 列」被容忍的原因：材料列缺失会取到 `""`，随后被材料校验拦截。

### m_Log / m_LogError
- 追加写 `gs_GetCommonAppDataDir() & "CDM_Import.log"`（该函数默认带尾反斜杠）。
- `m_Log` 内部 `On Error Resume Next`，日志失败绝不影响主流程。

---

## 五、g_RegenDoorLabelEMFs —— 重生成门板标签

> 用途：打开排版后的 ARD 嵌套图（**手动移动过门板**）后，逐件重新生成
> `<Job>_<材料>_<板名>_<件号>.emf`，覆盖旧文件并同步数据库。

### 流程

```
0.   校验排版档案：GetNestInformation 非空且 Sheets.Count > 0
     ⚠ VBA 的 Or 不短路 —— 必须拆成两次判断，否则 Ni=Nothing 时报错 91
0.5  【v1.7】弹「未保存修改」确认框（vbYesNo）；用户选"否"则直接退出
1.   初始化 COptions / strCTX（与排版时一致）
2.   从刀路属性 DEF_ATT_JOB_NAME 恢复订单名（取首个非空者）
3.   恢复材料名（四级回退）
4.   gbln_ConnectToDB() + m_PopulateNestingZones
5.   Set Material = New CMaterial : Material.MaterialName = sMat
5.0  ActiveDrawing.FullName → sUserARD（先存，供结尾/失败时重开）
     ActiveDrawing.SaveAs 到临时副本 regen_<Job>_<Timer>.ard
5.1  旧 <Job>_<材料>_*.emf 全部 Name 移到备份目录 regen_backup_<Job>_<材料>_<Timer>\
     （MkDir 失败则 sBakDir="" → 退化为不备份，不阻塞）
5.2  App.New → App.OpenDrawing 临时副本 → 【v1.7】blnScratchOpened=True
     构造临时嵌套路径 sScratchFile（v1.7 提前算好，供 EH 清理）
     m_CreateAlphaCAMDrawingsOfSheets Material, sNestOverride
     m_WaitForLabelEMFs(...) 轮询等待落盘
5.3  BeginTrans：预取本订单全部报表行，建两级索引
       · UID 索引   (DetailID, 板, PressPieceUID) → PK
       · 配对队列   (DetailID, 板) → 该组 PK（按 PK 升序）
     逐绘图实例：① 按 (DetailID, 板, UID) 精确命中
                 ② 无唯一码的旧行 → 从队列里取下一个未被认领的 PK，按序配对，并回填 UID
                 ③ 按 PK 写回 PressDoorImage + PressDoorCounter（缺列时 UID 单独跳过）
     → DELETE 本订单中已不存在的 DetailID
     → DELETE 本图各板中「未被认领」的行（多出来的重复行就是重复标签的来源）
     → CommitTrans
5.3b 成功 → RmDir 备份目录
5.4  删除临时嵌套 ard 与临时副本；ZoomAll；弹成功提示
结尾  App.New + App.OpenDrawing sUserARD（回到用户原档案）+ ZoomAll
EH   回滚事务 / 还原 EMF / 清理临时文件 / 重开原档案（见下）
```

> ### ⚠️ 标签行的身份 = 每件的 `PressPieceUID`（v1.9，2026-09-11）
>
> **现象**：多次调整板件位置后重新生成标签，会出现**重复标签**（两行指向同一个
> `..._Sheet A1_3.emf`）。
>
> **根因链**（已用现场数据证实）：
> 1. `PressDoorCounter` = `DEF_ATT_NEST_DOOR_COUNT`，是**板内实例序号**，
>    由 `Make.m_CreateAlphaCAMDrawingsOfSheets` 按 `SH.Parts` **枚举顺序**从 1 递增打下。
>    板件一被移动 / 重排，枚举顺序就可能变，**整板序号漂移**。
> 2. `AD_REPORT_DATA` 里存的是**旧序号**，只有整个报表流程重跑才会被重写。
> 3. v1.8 的 5.3 用 `DetailID + PressDoorCounter + SheetName` 匹配：
>    序号一漂移，本该配 `_2` 的那件**匹配不到任何行**，而 `_3` 那件**一次命中两行**
>    → 两行都写成 `_3.emf` → **重复标签**。
>
> **现场对照**（OrderID=10234 / `9-11纳百川` / Sheet A1，4 件）：
>
> | 绘图实例（`SH.Parts` 序） | DetailID | 绘图件序号 | 数据库旧值 | 数据库旧图 |
> |---|---|---|---|---|
> | 1 | 472438 | 1 | 1 | `_1.emf` |
> | 2 | **472439** | **2** | **3** ❌ | `_3.emf` ❌ |
> | 3 | **472439** | 3 | 3 | `_3.emf` |
> | 4 | 472437 | 4 | 4 | `_4.emf` |
>
> **没有一行是 2、却有两行是 3** —— 撞在一起的正是同一板件号（`PA锁Y`，数量 2）
> 的两件，也就是会被**框选一起移动**的那一对。
>
> **修法（A+B）**：
> - **A 稳定唯一码**：`Make.bas` 打件序号时同时写 `DEF_ATT_PIECE_UID`，**仅在缺失时分配**、
>   板内唯一（`U0001`…），移动板件后不变。5.3 优先按 `(DetailID, 板, UID)` 精确命中。
>   （已实测：该自定义属性名**可写可读可清空**，54 条路径全部成功。）
> - **B 按序配对 + 复位 `lngPK`**：没有唯一码的旧行，按 `(DetailID, 板)` 分组、
>   以 **PK 升序**与绘图实例**按序配对** → 数学上保证一一对应，两件不可能抢同一行；
>   同时把唯一码回填进该行。另外修复 `Make.m_InsertReportDataRouter` 里
>   `lngPK` **在 `For Each Ni` 循环中从不复位**（某件没匹配到行时会沿用上一件的 PK、
>   把上一件的行重复写一遍）的隐患。
> - 未被认领的重复行在事务内 **DELETE**。
>
> **要点：板件号（`DetailID`）会重复、件序号（`PressDoorCounter`）会漂移，
> 两者都不能单独当身份用；身份必须"每件一码"。**

### 材料名四级回退（步骤 3）

| 优先级 | 来源 | 说明 |
|---|---|---|
| 1 | 图纸名 `<Job>_<材料>.ard` 解析 | 最高优先，从 `ActiveDrawing.Name` 在 JobName 之后截取 |
| 2 | `AD_REPORT_DATA.PressDoorImage` 路径 | `SELECT TOP 1 ... WHERE INSTR(PressDoorImage, JobName) > 0`，在全路径中 `InStr` 定位 JobName 后取一段、再截到首个 `_` |
| 3 | `AD_ORDER_DETAILS.Material` | 单材料订单最可靠；`MoveNext` 后仍非 EOF 说明多材料 → 放弃 |
| 4 | 图纸属性 `DEF_ATT_MATERIAL_NAME` | 最后兜底 |

> ⚠️ **关键认知**：嵌套板 `MaterialName` 是 **SheetName 配置名**（如 "Admin"），
> **不是材料名**，所以不能直接从 Nest Part 取材料。
> 代码注释同时说明：**不用 `InStrRev` 切目录**（AlphaCAM VBA 中行为异常）。

### v1.7 失败现场还原（EH）

原实现失败后只回滚数据库 + 还原 EMF，会把用户留在空白图或临时副本里，
并在 ProgramData 残留 `regen_*` 临时文件。v1.7 补齐为四步：

1. **数据库**：`blnInTrans` 为真则 `RollbackTrans`
2. **标签文件**：`m_RestoreLabelBackup sBakDir` —— 把备份目录里的旧 EMF 全部 `Name` 回图片目录
3. **临时文件**：删除临时嵌套 ard（`sScratchFile`）与临时副本（`sTmpBak`）
4. **现场**：**仅当 `blnScratchOpened` 为真**（即 `App.New` 已把用户图撤下）时才
   `App.New` + `App.OpenDrawing sUserARD` + `ZoomAll` —— 否则重开会把用户
   **尚未保存的修改**一并冲掉

最后写日志并把「已回滚 / 已还原 / 已重开原档案」写进错误提示。

### 为什么 5.0 要 SaveAs 临时副本

`m_CreateAlphaCAMDrawingsOfSheets` 内部会打开/拆分/高亮图纸，
在**副本**上跑可保证用户正式图纸数据零触碰，保住**加工道次关联**；
这也解释了 3.4 版之前对主图做 `Name`/`Kill`（park/restore）会闪退的原因。

---

## 六、m_WaitForLabelEMFs —— 等待口径（v1.7 修正）

```
标签文件两类（都由 Make.m_ExportDoorLabelEMFs 产出）：
  整板图 <Job>_<材料>_<板名>.emf            每板 1 张   （Make.bas:3941）
  逐件图 <Job>_<材料>_<板名>_<件号>.emf     每板 = 该板 SH.Parts.Count 张（Make.bas:3944）
```

该函数用 `Dir$(sDir & "*" & DEF_EXTENSION_EMF)` 统计 `<Job>_<材料>_*.emf`，
**两类文件都会被计入**。因此期望值必须是：

```
lngExpected = Σ Nsh.Parts.Count (总件数) + 板数
```

- **v1.6 及以前**只传总件数 → 整板图被误计入，条件 `n >= expected` 可能**提前满足**。
- **v1.7** 在调用点传 `lngExpectedDoors + lngSheetCount`。

> 实现上仍以「目录计数」而非「按文件名逐个点名校验」，是为了避免依赖
> `Nsh.Name` 与 `Make.bas` 里 `colSheetNames(iCount)` 的字符串一致性（未在实机验证）。

---

## 七、外部依赖（不可独立编译）

本模块**不能脱离 CDM 工程单独编译**，大量符号来自 `Make.bas`(284KB) / `Events.bas`(106KB) /
`Functions.bas` / `Globals.bas`：

| 类别 | 符号 |
|---|---|
| 生产 | `g_Make_Master`、`m_CreateAlphaCAMDrawingsOfSheets`、`m_PopulateNestingZones` |
| 数据库 | `gbln_ConnectToDB`、`gdb_CDM`（`ADODB.Connection`，故支持 `BeginTrans` 等）、`gs_FixSQL`（`Replace$(s,"'","''")`）、`gvar_CheckNull` |
| UI | `Frame.ShowProgressBox` / `CloseProgressBox`、`App.Frame.WindowHandle` |
| 对象 | `CMaterial`、`COptions`（`clsOptions.CTXFile` / `PathToRoot` / `OutputResultsSubFolder`）、`strCTX` |
| 常量 | `DEF_ATT_JOB_NAME`、`DEF_ATT_MATERIAL_NAME`、`DEF_ATT_DETAIL_ID`、`DEF_ATT_NEST_DOOR_COUNT`、`DEF_PATH_IMAGE`、`DEF_EXTENSION_EMF`/`DEF_EXTENSION_ARD`、`DEF_UNDERSCORE`、`DEF_BACKSLASH` |
| 工具 | `gstr_EnsureBackslash`（幂等）、`gstr_CheckDir`、`gs_GetCommonAppDataDir`（默认带尾反斜杠） |
| 排版对象 | `NestInformation` / `NestSheet` / `NestPartInstance` |
| 全局 | `gstr_JobName` |
| 窗体 | `frmAutoNest`（`AutoImportNest` 弹出，`AutoImportNestWithParams` 由它调用） |

菜单链路：`CDM功能/Events.bas:277` 注册菜单项 → `Events.bas:2857 m_AutoImportNest()` →
`modAutoImportNest.AutoImportNest` → `frmAutoNest.Show vbModeless`。

---

## 八、版本演进

| 版本 | 日期 | 要点 |
|---|---|---|
| v1.0 | — | 初始：弹窗选 CSV → 导入 → `g_Make_Master` |
| v1.1 | — | 日志落盘、错误带步骤名、失败行明细、返回码常量、清 `Nest Completed` 残留键 |
| v1.3 | — | 重生成传临时嵌套路径，绝不 `SaveAs` 覆盖用户主图（保留加工道次） |
| v1.5 | — | 删除对主图 `.bak` 驻留/还原（`Name`/`Kill`，**闪退根源**）；结尾重开真档案 |
| v1.6 | 2026-09-02 | 旧 EMF 先备份→成功删除→失败回滚；DB 同步加事务；去掉 `Exit For` 漏更；标签导出抽为 `Make.m_ExportDoorLabelEMFs`；材料改为只校验（`m_CheckMaterialExists`） |
| **v1.7** | **2026-09-10** | **ImportCSV 全流程事务化**；**重生成失败现场还原**（重开原档案 + 清临时文件）；**未保存修改确认框**；CSV 存在性检查上移；删死代码 `sDefaultMaterial` / `glng_EnsureMaterial`；`Sleep` 补 `PtrSafe`；**`m_WaitForLabelEMFs` 期望值改为「板数+总件数」** |
| **v1.8** | **2026-09-10** | **修屏幕刷新泄漏（B1）**：配合 `Make.bas:3991/3992` 恢复 `ScreenUpdating` / `ProjectBarUpdating`，并在本模块成功与失败路径兜底 `Redraw` —— 解决"窗口标题停在 `regen_<订单>_<Timer>` 临时副本名"；**备份目录改 Kill+RmDir（B2）**，原先 `RmDir` 对非空目录必失败、累积 28 个；**临时嵌套 ard 改在 `App.New` 关档后补删并记日志（B3）** |

---

## 九、已知遗留问题

> ### ✅ v1.8 已修（2026-09-10 实机定位并验证）
>
> **B1 屏幕刷新泄漏 —— "窗口还显示临时档案"的真正原因**
> `Make.m_CreateAlphaCAMDrawingsOfSheets` 在 `Make.bas:3742/3743` 置
> `ActiveDrawing.ScreenUpdating = False` 与 `Frame.ProjectBarUpdating = False`，
> 而收尾处的恢复语句 `3991/3992` **被注释掉了**（同段的 `App.DisableUndo`(3993)、
> `QuickShading`(3995) 都有配对恢复，唯独这两项漏了）。
> 后果：每次生产排版 / 重生成标签之后 AlphaCAM 不再重绘，窗口标题与画面停在旧的
> `regen_<订单>_<Timer>` 上，**看起来像"还打开着临时档案"，其实 `ActiveDrawing.FullName`
> 早已是真档案、`Modified=False`**。
> 已恢复 Make.bas 两行，并在本模块成功路径与 EH 都兜底 `ScreenUpdating=True` + `Redraw`。
>
> **B2 备份目录永远删不掉**
> 备份目录里装的是 5.1 移走的旧 EMF，**非空目录 `RmDir` 必定失败**，且被
> `On Error Resume Next` 吞掉 → ProgramData 累积 **28 个** `regen_backup_*`（222 个旧 EMF，
> 时间跨度 9/2–9/10）。已改为先 `Kill sBakDir & "*.*"` 再 `RmDir`。
>
> **B3 临时嵌套 ard 删不掉**
> 5.4 删除时该文件仍被 AlphaCAM 占用（`m_Create` 打开过它），失败被静默吞掉 →
> 累积 `regen_scratch_*.ard`。已改为在 `App.New` 关档之后补删一次，并把结果写入
> `CDM_Import.log`（`临时嵌套档案已清理` / `仍未能删除(被占用)`）。
> 实测：关档后该文件 `locked=False`，可正常删除 —— 证实了占用判断。

以下为 v1.7 时点仍存在的问题：

1. **材料回退优先级 3 的子查询假定 JobName 唯一**：
   `SELECT DISTINCT Material FROM AD_ORDER_DETAILS WHERE OrderID=(SELECT OrderID FROM AD_ORDERS WHERE JobName='…')`
   —— 若 `AD_ORDERS` 中存在同名订单，子查询返回多行会报错，被外层
   `On Error Resume Next` 静默吞掉后退到优先级 4。建议改为 `IN (SELECT …)`。
   （优先级 2 的 `PressDoorImage` 查询已带 `TOP 1`，无此问题。）
2. **`Nsh.Name` 与 `colSheetNames` 的一致性未验证**：5.3 写 `PressDoorImage` 用的是
   `Nsh.Name`，而 `m_ExportDoorLabelEMFs` 收到的是 `Make.bas` 的 `colSheetNames(iCount)`。
   若两者不等，DB 中的图片路径会与实际文件名不符（等待逻辑已刻意规避该依赖）。
3. **数值字段直接拼接进 SQL**（`q`/`w`/`h`/`lngStyleNum`/`Val(sGrp)`/`lngOrderID`）：
   依赖上游 `Val()` 保证安全；文本字段一律经 `gs_FixSQL`。
4. **`AutoImportNestWithParams` 的 EH 无条件调用 `Frame.CloseProgressBox`**：
   在「解析CSV」「连接数据库」阶段出错时进度框并未打开。
5. **`Sleep` 用 `Declare PtrSafe`**（与 `frmAutoNest.txt` 一致，假定 VBA7）；
   若移植到 VBA6 宿主需加 `#If VBA7` 条件编译。
6. **未落地**（见 `标签重生成优化方案.md`）：方案 C 增量重生成、D.2 材料名由入口直接传参、
   D.4 临时文件 `GetTempName`/GUID 独立子目录、逐件 `SaveEmfFile` 结果校验。

---

## 十、验证要点（部署后回归）

1. **正常导入**：CSV 全字段齐全 → 订单/明细正确落库，`bRunNest=True` 走完生产+排版+NC。
2. **材料缺失**：故意写不存在的材料 → 整单失败、**数据库无任何新增**（事务回滚）、提示清晰。
3. **覆盖模式中途失败**：勾选强制覆盖 + 构造第 N 行失败 → 确认**旧订单数据仍在**（回滚生效）。
4. **文件路径错误**：填不存在的 CSV → 确认**不再产生空订单/新客户**。
5. **只导入订单**：勾选后不调用 `g_Make_Master`，仅落库。
6. **重生成标签（成功）**：打开 ARD → 手动移板 → 重生成 → 标签更新、DB 同步、结尾回到原档案。
7. **重生成标签（失败）**：中途制造错误（如只读图片目录）→ 旧标签完好、当前图=原档案、无 `regen_*` 残留。
8. **未保存修改**：改图不保存 → 点重生成 → 确认框选"否"应无任何副作用。
9. **多板 / 多材料 / 订单名含 `_`** 边界样例。
