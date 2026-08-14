# Make.bas 分段详细分析报告

> 对象：`CDM功能/Make.bas`（CDM 工程加工引擎源码备份）
> 规模：290,252 字节 / 7,926 行 / 71 个过程 / 纯 ASCII（英文代码+英文注释）
> 编码：GBK（本文件实际为纯 ASCII）

## 一、模块结构总览

| 段 | 行区间 | 行数 | 主要内容 |
|----|--------|------|----------|
| S1 | 1–886 | 886 | 嵌套排版基础：删嵌套板、自定义宏、手柄孔、排版区域、NC 输出、打包到 LHS |
| S2 | 886–1643 | 758 | 排版区域集合、规则(Rules)、刀具排序、压机总控 `gbln_Make_Master_Press`、条码 |
| S3 | 1643–2274 | 632 | 压机排版 `m_DoNestingPress`、路由器排版 `m_DoNestingRouter`、对比、MPR 保存 |
| S4 | 2274–2811 | 538 | ⭐ **主入口 `g_Make_Master(strOrder)`** |
| S5 | 2811–3632 | 822 | 报表数据填充（Press/Router）、文件名编译、预览 `g_Preview_Master` |
| S6 | 3632–4591 | 960 | 创建 AlphaCAM 图纸（Router/Press）、零件处理、门/材料数据填充、输出、ARD 插入 |
| S7 | 4591–5773 | 1183 | 初始化、路径属性、路径数据、后处理变量、加工执行、更新保存、`m_SmallFirst` |
| S8 | 5773–6990 | 1218 | 填充件、报表数据插入、订单信息、嵌套拆分、PGM 转换、偏移加工 |
| S9 | 6990–7926 | 937 | 偏移/无偏移加工、属性设置、门型 Style 900/910/920/930 加工 |

## 二、过程索引（71 个）

| 行号 | 过程 | 可见性 | 段 |
|------|------|--------|----|
| 4 | `m_DeleteNestSheet(NestSheetToDelete As NestSheet)` | Private | S1 |
| 36 | `mbln_RunCustomMacro(MacroFilename, Door) As Boolean` | Private | S1 |
| 130 | `mbln_DrawHandleHole(Door, HoleSize, HoleXPos, HoleYPos) As Boolean` | Private | S1 |
| 142 | `mbln_DrawHandleHoles(Door, bPreview) As Boolean` | Private | S1 |
| 342 | `m_DrawNestZones(Material, ZonesCollection)` | Private | S1 |
| 373 | `m_NCHopsOutput(SheetLength, ...)` | Public | S1 |
| 505 | `m_NCHopsOutputTemp()` | Public | S1 |
| 567 | `m_PackFinalNestSheetToLHS(NestMaterial As CMaterial)` | Private | S1 |
| 886 | `m_PopulateNestingZones()` | Public | S2 |
| 922 | `m_PopulateNestZoneCollection(rstNestingZones, TargetCollection)` | Private | S2 |
| 959 | `m_PopulateRulesCollection(CurrentRouter As CRouter)` | Private | S2 |
| 1019 | `mbln_RulesApply(Door As CDoor) As Boolean` | Private | S2 |
| 1077 | `m_RuleTestValid(Door, TestItem, TestOperatorID, TestValue) As AdoorRuleTestResult` | Private | S2 |
| 1154 | `m_ToolSorting()` | Private | S2 |
| 1170 | `mbln_ToolOrdering() As Boolean` | Private | S2 |
| 1307 | `gbln_Make_Master_Press() As Boolean` | Public | S2 |
| 1608 | `gstr_GenerateNestedSheetBarcode() As String` | Public | S2 |
| 1643 | `m_DoNestingPress(Press, PressColour, Optional PressThickness)` | Private | S3 |
| 1792 | `m_DoNestingRouter(Router, Material, colNestANC)` | Private | S3 |
| 2162 | `mbln_CompareRouterNestToOrder(Router, MaterialName) As Boolean` | Private | S3 |
| 2216 | `m_MprSave(MprLength, MprWidth, ...)` | Private | S3 |
| 2274 | ⭐ `g_Make_Master(strOrder As String)` | Public | S4 |
| 2811 | `m_PopulatePressData(OrderString)` | Public | S5 |
| 2925 | `m_PopulateRouterData(OrderString)` | Public | S5 |
| 3004 | `mbln_PopulateReportData(OrderString) As Boolean` | Public | S5 |
| 3080 | `mbln_PopulateReportDataPress(OrderString) As Boolean` | Public | S5 |
| 3144 | `mstr_CompileNestOrNCFilename(Material) As String` | Private | S5 |
| 3174 | `mstr_CompileNestOrNCFilenamePress(PressDetails) As String` | Private | S5 |
| 3198 | `mlng_GetCIM_INCID_Value(Door) As Long` | Private | S5 |
| 3219 | `mstr_GetCustomMacroName(MacroFilename) As String` | Private | S5 |
| 3250 | `mstr_GetJobName(OrderID) As String` | Public | S5 |
| 3267 | `g_Preview_Master(lType)` | Public | S5 |
| 3406 | `m_CreateAlphaCAMDrawingsOfSheetsPress(PressDetails)` | Public | S6 |
| 3632 | `m_CreateAlphaCAMDrawingsOfSheets(Material)` | Public | S6 |
| 3997 | `mbln_ComparePressNestToOrder(Press, PressColour, ...) As Boolean` | Private | S6 |
| 4030 | `mbln_ProcessPart(Door, lngPartNumber, ...)` | Private | S6 |
| 4148 | `m_FillDoorData(Door, rDetails, bPreview)` | Private | S6 |
| 4219 | `m_FillMaterialData(Material, rstMaterial)` | Private | S6 |
| 4274 | `m_DeleteReportData(lOrder, sJob)` | Private | S6 |
| 4314 | `mstr_CurrentPost() As String` | Private | S6 |
| 4323 | `m_OutputToAlphaEDIT(colNestANC, colDOORANC)` | Private | S6 |
| 4372 | `m_UpdateDBNestPaths(sFieldName, sFieldPath, ...)` | Private | S6 |
| 4435 | `mbln_InsertARD(Door, bPreview) As Boolean` | Private | S6 |
| 4591 | `mbln_Intialize(lngGeo) As Boolean` | Private | S7 |
| 4618 | `m_SetAttributes(Pth, Door)` | Private | S7 |
| 4661 | `mbln_FillPathData(rst) As Boolean` | Private | S7 |
| 4799 | `m_SetPostVariables(dPanel_X, dPanel_Y, ...)` | Private | S7 |
| 4824 | `mbln_MakeMachining(Door, lngGeoNumber, lngPartNumber, ...)` | Private | S7 |
| 5139 | `mbln_UpdateAndSave(Door, lPartNumber, ...)` | Private | S7 |
| 5585 | `mbln_UpdateAndSavePress(Door, lPartNumber) As Boolean` | Private | S7 |
| 5773 | `m_SmallFirst()` | Public | S7 |
| 5788 | `mbln_InsertFillerParts(Material) As Boolean` | Private | S8 |
| 6015 | `m_InsertReportDataPress(Press, PressColour, Optional PressThickness)` | Private | S8 |
| 6168 | `m_InsertReportDataRouter(Material)` | Private | S8 |
| 6461 | `m_FillOrderInfo(DoorComponent, rOrder, sJob, lJob)` | Private | S8 |
| 6515 | `m_SplitNestANC(Material, nstThisNest, ...)` | Private | S8 |
| 6751 | `mbln_HasSTART(strThisFile) As Boolean` | Private | S8 |
| 6798 | `m_ConvertToPGM(strFile)` | Private | S8 |
| 6839 | `m_CheckForErrors(strFilename)` | Private | S8 |
| 6899 | `mpths_PathsInGroup(pPaths, iGroup) As Paths` | Private | S8 |
| 6930 | `mpths_PathsNotInGroup(pPaths, iOffsetFrom) As Paths` | Private | S8 |
| 6964 | `mbln_MachineWithOffset(cPath, pthPath, pthPick, ...)` | Private | S8 |
| 7190 | `mbln_MachineWithNoOffset(cPath, pthPick, ...)` | Private | S9 |
| 7252 | `m_SetDetailAttributes(Door, P)` | Private | S9 |
| 7300 | `m_SetPressAttributes(Door, P)` | Private | S9 |
| 7320 | `mbln_Style_900(Door, lngPartNumber, colANC, bPreview, bPress, bStopProcessing)` | Private | S9 |
| 7467 | `mbln_Style_910(Door, lngPartNumber, colANC, bPreview)` | Private | S9 |
| 7577 | `mbln_Style_920(Door, lngPartNumber, colANC, bPreview)` | Private | S9 |
| 7692 | `mbln_Style_Make_930(Door, bPress, Optional lngPartNumber, ...)` | Private | S9 |
| 7901 | `Temp()` | Public | S9 |
| 7916 | `TestScrap()` | Public | S9 |

<!-- 以下为分段详细分析（S1–S9），逐段补充 -->

## S1 段分析（1–886 行）：嵌套排版基础工具

### m_DeleteNestSheet（4–34）
删除指定嵌套板及其全部组成：选中板内所有零件路径（`Npi.Paths`）、板路径本身、板文本（`DEF_ATT_SHEET_ITEM_NOS="1"` 且 `DEF_ATT_SHEET_IDENT` 匹配）、泡形几何（`DEF_ATT_SHEET_BOBBLE`），最后 `ActiveDrawing.DeleteSelected` 一次删除。用于排版结果回滚。

### mbln_RunCustomMacro（36–128）— ⭐ 930 用户门型宏的核心
1. 若 `gstr_CustomMacroProjectName` 为空或文件名变化 → `App.LoadAddIn` 加载宏工程，用 `mstr_GetCustomMacroName` 取工程名
2. 将 `Door.UserVariableString` 按 `";"` 拆分进 `UserVariablesCollection`（`PDbl` 转数值），`UserVariableDescriptionString` 拆分进描述集合——即 modAutoImportNest 导入时复制的 `UserVariableString/UserValue_0~6` 在此消费
3. 工程未打开时再次 `LoadAddIn`；`Door.JobName = gstr_JobName` 供宏使用
4. `App.Run(工程名.DEF_CUSTOM_MACRO_MODULE_NAME.DEF_CUSTOM_MACRO_PROCEDURE_NAME, Door, UserArg_0..6)` 调用宏
5. 结束后若工程是本次加载的则 `App.EnableAddIn ..., False` 卸载；`Door.CustomMacroSuccess` 为 False 判定失败
6. 错误 -2147467259 → 显示定义用户门型帮助；其余写 `WriteError`

### 手柄孔（130–340）
- `mbln_DrawHandleHole`：范围内画圆（`CreateCircle`）并标 `DEF_ATT_HANDLE_GEO="1"`
- `mbln_DrawHandleHoles`：读取 `grst_GetHandleConfiguration(Door.HandleID)` 配置 → 按 9 种基准位置（左下/下中/右下/左中/居中/右中/左上/上中/右上）× 孔方向（垂直/水平）× 数量与间距计算孔位（偶数孔围绕基准均布）→ 逐孔画圆 → 应用 `MillMachiningStyles(LicomdirPath & MachiningStyle)` 加工样式生成钻孔刀路并标 `DEF_ATT_HANDLE_TP` → 孔位越界时预览版直接提示、生产版询问是否继续
- 无 HandleID（=0）直接跳过

### m_DrawNestZones（342–368）
按 `Door.NestingZone` 画排版区域矩形（自动区域仅在 `AutoZoneInUse` 时画、手动区域必画；每区域只画一次，`ZoneDrawn` 防重），收集进 `pthsNestZones`。

### NC 输出（373–565）
- `m_NCHopsOutput`：逐张嵌套板图纸——第 2 块起整体平移到 (0,0)（按 NestInformation 左下角）→ 保存 → 调外部工程 `NcHopsPost.Make.CallFromExternal`，传 `Length;Width;Thickness;XShift;YShift;ZShift;Comment;b5Axis;bAcamToolValues;Language;strNCName` 生成 `.hop` 文件。偏移量来自注册表 `NCHops_OffsetX/Y/Z`，Z 默认 `-厚度`；无 NcHopsPost 工程时提示
- `m_NCHopsOutputTemp`：硬编码测试版（1200×2000×16 → `C:\hopstest.hop`）

### m_PackFinalNestSheetToLHS（567–883）— 最后一块板收紧优化
对最后一块嵌套板（LHS 左侧）做双重验证排版：
1. 收集板上第一个刀路零件（`DEF_ATT_ARD` 文件名 + `DEF_ATT_DOOR_ROTATION_ANGLE`）到两个 CNest：矩形算法 `StartNestListRectRouter` + 真形算法 `StartNestListTrueRouter`
2. 各自新建板（尺寸 0.18*SheetLength 偏移）→ `N.Nest` 排版
3. 结果校验：`NestErrorsNl.Count>0` → 报未嵌套零件清单并终止；产出版>1 → 回滚删除并清掉 `LicomUKja_cutline_mark` 残留
4. 正常时比较两方案"最左刀路 extent"（`dblLeftmostTP*`），保留更紧凑者、删除另一方案的板；两方案都失败则保留原始板
5. 所有对象置 `Nothing` 清理

**与 modAutoImportNest 的关系**：本段运行自定义宏（930 门型几何生成）依赖导入时正确写入 `UserVariableString/UserValue_0~6`；`mbln_RunCustomMacro` 即 README 所述"无法连接用户定义的宏"报错的发生处（`-2147467259`）。


## S2 段分析（886–1643 行）：排版区域 / 规则 / 刀具排序 / 压机总控

### 排版区域（886–956）
- `m_PopulateNestingZones`：清空全局集合 `colManualNestZones/colAutoNestZones`，重建 `pthsNestZones` 路径集合；分别按 `adoorNestZone_Manual` / `_Automatic` 从数据库读取区域
- `m_PopulateNestZoneCollection`：recordset → `CNestingZone` 对象（ZoneID/Name/Number/起止坐标/宽高/PartDimension/PartArea/AssignMethod），加入集合，键 `"k" & ZoneID`

### 规则系统（959–1151）
- `m_PopulateRulesCollection`：按路由器后处理器文件名（`gstr_StripLicomDatPath`）从数据库 `grst_GetRulesForRouter` 读规则到 `colRouterRules`（RuleID/OperatorID/测试变量/阈值/结果后处理器/结果路由器/文本）
- `m_RuleTestValid`：测试门宽或门高，支持 =、>、<、>=、<=、<> 六种运算符
- `mbln_RulesApply`：逐条规则判定；命中时按 `ResultRouter` 设置门旋转方式 `RotationMethod`（`adoorRULE_RESULT_ROTATE_LOCK→LOCKX`、`LOCK_ONLY→LOCKY`、`FREE_ROTATE→FREE`）；测试出错则弹规则文本并返回失败

### 刀具排序（1154–1304）
- `m_ToolSorting`：选项 `ToolOrderingUsed` 开启时，在 `SuppressUpdateRapids` 包裹下执行 `mbln_ToolOrdering`
- `mbln_ToolOrdering`：按数据库 `grst_GetToolOrder` 定义的刀具顺序（ToolName/Number/OffsetNumber）重排操作——
  1. 逐嵌套板（**排除 `_REV` 反面板**）取首末操作号
  2. 按刀具序列表逐条扫描该板操作，工具三要素匹配 → `Operations.Renumber(当前, 插入点, acamOpINSERT_IN_FRONT)`
  3. 性能优化（TFS#57996）：复用 `Ops` 引用而非反复访问 `ActiveDrawing.Operations`
  4. `App.DisableUndo = True` 包裹防撤销栈膨胀

### ⭐ gbln_Make_Master_Press（1307–1606）— 压机排版总控
- 三层遍历：`colPressData`（压机）→ `PressColour`（颜色/膜）→ `PressThickness`（厚度）→ `PressDoor`（门板）
- 每颜色新建图纸（`App.New`），`mbln_ProcessPart(PressDoor, ..., bPress=True, ...)` 逐门加工；失败 → 注册表 `COMPLETENEST=0` + 提示 + 退出
- 按选项分组排版：`GroupByMaterialThickness=True` → 每个厚度一组调用 `m_DoNestingPress(Press, PressColour, 厚度)`；否则按颜色整体一组调用 `m_DoNestingPress(Press, PressColour)`
- ⚠️ **1404–1603 为约 200 行整块注释掉的旧版嵌套代码**（死代码，保留历史）

### gstr_GenerateNestedSheetBarcode（1608–1641）
`Barcode.ini`（`g_MoveBarcodeIni` 定位）读取 `SheetID/NextSheetID` 并自增返回，用作嵌套板条码号。


## S3 段分析（1643–2274 行）：压机/路由器排版执行、校验、MPR

### m_DoNestingPress（1643–1790）— 压机排版执行
1. 清空嵌套列表 → 加载 `clsNest.NestListName`（可选 `Nl.OrderParts` 大件优先编号）
2. 建板矩形（SheetWidth×SheetLength）→ `N.Nest` 执行排版
3. 错误：`NestErrorsNl.Count>0` → 列出未嵌套件（数量×类型（宽*长））并退出；否则 `mbln_ComparePressNestToOrder` 数量校验
4. 给每块板路径标属性：`DEF_ATT_PRESS_NAME` / `DEF_ATT_FOIL_COLOUR` / `DEF_ATT_SHEET_THICKNESS`
5. 按选项保存：不保存 ANL 则删嵌套列表文件；`OutputToSecondLocation` 时保存副本（可建子文件夹）；`SaveAllNestARD` 保存图纸（文件名 `PressName_ColourName[厚度].ard`）
6. 报表：`m_InsertReportDataPress`（`DisableReports` 关闭时跳过）

### ⭐ m_DoNestingRouter（1792–2160）— 路由器排版执行（主路径）
压机版的全功能超集，完整流水线：
1. 嵌套准备：加载列表（可选大件优先）→ **边角料嵌套 `CreateObject("StdAlpha.ShareClass").ScrapNesting(...)`**（DLL 扩展，`EditMark` 标注的第三方钩子）→ 画自动/手动排版区域
2. **Twin Head 双头嵌套**：设 `DEF_ATT_CAN_START_TWIN_HEAD="1"` + `DEF_ATT_TWINHEAD_ORDERSTRING=gstr_JobIDs`，清注册表标志，`SuppressUpdateRapids False`（供 NestComplete 事件宏）
3. 执行 `N.Nest` → 未嵌套错误清单 / `mbln_CompareRouterNestToOrder` 数量校验
4. 后处理链（按选项开关）：
   - `PackFinalSheetToLHS` → `m_PackFinalNestSheetToLHS`（S1 收紧优化）
   - `InsertFiller` → 禁用 TwinHead + `mbln_InsertFillerParts`（失败提示）
   - `CutSmallFirst` → `m_SmallFirst`；`OnionSkin` → `g_DoOnionSkin`
   - `ProcessWaste` → `ProcessWasteMaterial`（废料清理策略/间距）
   - `m_ToolSorting` 刀具排序
5. **报表**：TwinHead 标志≠"1" 时才 `m_InsertReportDataRouter`（双头自带报表防重复）
6. **反面嵌套**：设 `DEF_ATT_CAN_START_REVERSE_NEST` → 若标志为 1 调 `AlphaDOOR_ReverseSideNesting.Main.ReverseSideNestingMain` → 清标志
7. 程序输出：`m_SetPostVariables` → 文件名 `mstr_CompileNestOrNCFilename` → 标 `DEF_ATT_ALPHADOOR/NC_FILE_EXTENSION` → 双位置输出（MPR / NCHops / NC 三种模式）→ `SplitNestedPrograms` 时 `m_SplitNestANC` 拆分 → `m_UpdateDBNestPaths` 更新 DB 路径 → 板属性标 ANC 名 → 保存 ARD → 顺序编号（`UseOrderName=False` 时 `gbln_IncrementNCSeqNum`）

### mbln_CompareRouterNestToOrder（2162–2212）
验证"嵌套零件数 = 订单数量"：累加 `colDoors` 的 `Quantity`（**排除 `ByPassNest` 门**，TFS#49412）；对比嵌套板 `Parts.Count`，多头（`DEF_ATT_NUM_HEADS`）与 TwinHead 路径（`DEF_ATT_TWIN_HEAD_PATH="1"` 计 ×2）加权。

### m_MprSave（2216–2270）
调用外部 WoodWop 工程 `WoodWop4xV20_edge/normal.Make.MprFromExternal` 生成 MPR 文件，13 个分号分隔参数（长宽厚/已加工件/未加工件/废料偏移/注释/文件名），偏移值取自 `clsOptions.MPR_*`。


## S4 段分析（2274–2811 行）：⭐ 主入口 g_Make_Master(strOrder)

### 完整执行流水线

**1. 初始化与准备**
- 注册表 `COMPLETENEST=0`（先假设失败——modAutoImportNest 读取此值判断成败）
- 记录当前后处理器 `mstr_CurrentPost`，关 3D 视图，`g_LockAcam` 锁界面
- 兼容补丁：Acam 2011 群组号丢失 bug → 强制 `DisableUndo=False`（Salesforce 5002000000DYiqB / TFS41956）
- 实例化 `clsOptions/clsPathData/clsTypeData`，建集合；`colVBAUserStyles` 空时 `g_GetVBAProjects` 加载用户样式工程
- **多订单支持**：`strOrder` 逗号分隔 → `JobName="MergedOrder"`、CustomerID=1；单订单 → `mstr_GetJobName` + `glng_GetCustomerIDFromOrderID`

**2. 环境准备**
- 进度框显示；创建输出目录（metafile / 订单子文件夹 / 第二位置子文件夹）
- `gbln_ConnectToDB` 连库 → `m_PopulateNestingZones`（排版区域）→ `m_PopulatePressData`（压机数据）→ `m_PopulateRouterData`（路由器数据）
- 逐订单 `m_DeleteReportData` 清空旧报表数据

**3. 压机支线**
`colPressData.Count>0` 时：`mbln_PopulateReportDataPress` → `gbln_Make_Master_Press`（S2 压机排版）→ `mbln_PopulateReportData`

**4. 路由器主线（核心循环）**
```
For Each Router In colRouter
  选后处理器（ONEPOST 跳过；VBA post 备份/覆写 APC 设置；DefaultPost 缺失→提示用当前 post）
  m_PopulateRulesCollection Router     ← 规则
  For Each Material In Router.colMaterials
    App.New → 设材料 → StartNestListRouter
    For Each Door In Material.colDoors
      NestingOption=DISABLED → Door.ByPassNest=True
      mbln_ProcessPart(Door, lngPartNumber, colDOORANC, False, False)   ← 加工单门
       失败 → COMPLETENEST=0 + 提示 + 退出
       成功 → Alphacim 集成：mlng_GetCIM_INCID_Value + gbln_UpdateAlphacimComponentGrouping
      ByPassNest 之外的门 → blnNestingExists=True
    Next
    blnNestingExists → m_DoNestingRouter(Router, Material, colNestANC)   ← S3 排版执行
  Next
Next
```

**5. 压机报表数量回填**
遍历 `colPressReportData`，按 `DetailID_SheetName_RouterSheetName` 去重聚合 `CPressReportQty`，逐条 `UPDATE AD_REPORT_DATA SET PressQuantityThisSheet=?`——压机与路由器交叉件数量回写（2661 行前，2619 行起为独立代码块）

**6. 收尾**
- 有零件：非 `OutputToNCOnly` → `m_OutputToAlphaEDIT`（AlphaEDIT 输出）；`COMPLETENEST=-1` 成功标志；**询问是否继续下一个订单**（`CREATEANOTHER` 注册表）；更新订单 `ProcessedDate`
- 无零件：`COMPLETENEST=0` + "nothing done" 提示
- `Controlled_Exit`：恢复后处理器与 APC 设置 → `KeepLastDrawingOpen` 选项控制最后图纸去留 → 删除 `colDeleteFiles` 临时文件 → 释放对象 → `g_UnLoadAllForms` → `g_UnlockAcam`
- 错误：`COMPLETENEST=0` + MsgBox + `WriteError`

### 与 modAutoImportNest 的联动点
- `g_Make_Master(CStr(lngOrderID))` 是自动化生产排版的排版引擎；导入的订单/明细经 `m_PopulateRouterData/PressData` 进入 `colRouter.colMaterials.colDoors` 后由本过程驱动
- 成败判定 `GetSetting("LICOM AlphaDOOR","Nest Parameters","Nest Completed")` ← 即本过程的 `COMPLETENEST`（-1/0）


## S5 段分析（2811–3632 行）：数据装配 / 报表 / 命名 / 预览 / 单板图纸

### m_PopulatePressData（2811–2921）— 压机数据装配
从订单明细构建 `colPressData` 四层树：**压机 CPress**（PressID/名称/长宽/板边距/件距/NumberBySize/PartRotation/UseTrueShape/打包方向，`RectPackTo` 空值回退 2）→ **颜色 CPressColour** → **厚度 CPressThickness** → **门板 CDoor**（`m_FillDoorData` + `m_FillOrderInfo`）。全部按名称/厚度作集合键去重。

### m_PopulateRouterData（2925–3001）— 路由器数据装配 ⭐（导入数据入口）
按 **后处理器** 分组构建 `colRouter` → **材料 CMaterial**（`grst_GetMaterial` + `m_FillMaterialData`）→ **门板 CDoor**（`m_FillOrderInfo` + `m_FillDoorData`）。
- 明细无 PostProcessor → 回退 `clsOptions.DefaultPost`
- ⚠️ **modAutoImportNest 导入的 `AD_ORDER_DETAILS` 行正是经此过程进入 `colRouter.colMaterials.colDoors`**，随后被 `g_Make_Master` 消费

### 报表数据对象（3004–3135）
- `mbln_PopulateReportData`：按订单读 `grst_GetReportDetails` → `colReportData`（DetailID → ReportDataPK 列表）；无数据 → 提示 + False
- `mbln_PopulateReportDataPress`：按 `grst_GetOrderDetailsByID` 的 PK → `colPressReportData`

### 文件名编译（3144–3194）
- `mstr_CompileNestOrNCFilename`：`[前缀_][订单名|顺序号(glng_GetNCSeqNum)][_材料名|材料序号]`
- `mstr_CompileNestOrNCFilenamePress`：路由器版 + `_PressDetails`

### Alphacim 集成（3198–3217）
`mlng_GetCIM_INCID_Value`：按 `gstr_AlphacimPKField` 配置从门板字段（`CSV_CustomerName/CSV_OrderNumber/CSV_ItemNumber/ProductionComment/CustomField1/2`）取 Long 作为 INC_ID——**modAutoImportNest 导入的 CSV_* 字段在此被消费**。

### mstr_GetCustomMacroName（3219–3248）
按宏文件名在 `App.VBE.VBProjects` 中匹配工程名（`Project.FileName` 含文件名；未保存工程 Err 76 → 跳过），并缓存 `gstr_CustomMacroFileName`。

### g_Preview_Master（3267–3404）— 门型单件预览
按 `lType` 读门型数据 → `m_FillDoorData(bPreview=True)` → `mbln_ProcessPart(..., True)` 加工 → `UserSelectOneGeo` 循环等待用户点选/ESC → 清理恢复。

### m_CreateAlphaCAMDrawingsOfSheetsPress（3406–3630）— 压机单板图纸 + 报表 EMF
1. 保存总嵌套图纸 → 文本转几何（保留 `attSheetIdent` 属性）
2. 每块板：`CreateTempDrawing` 拆分为单板图纸，标 `DEF_ATT_POCKET_PATH`（ProcessType 2/4/5），保存 `JobName_PressDetails_板名.ard`
3. 报表图：临时关背景渐变（TFS#45271）、白底、黑色/黄色（槽刀路）着色、按板范围 `ZoomToBox` → `SaveEmfFile` 生成 EMF
4. 恢复背景与原始嵌套图纸


## S6 段分析（3632–4591 行）：单板图纸 / 零件分派 / 数据填充 / ARD 插入

### m_CreateAlphaCAMDrawingsOfSheets（3632–3995）— 路由器单板图纸 + 报表图（压机版的增强版）
1. 关 3D/快速着色/屏幕刷新，`DisableUndo`，隐藏排版区域（报表不含）
2. 每板每零件路径标 `DEF_ATT_NEST_DOOR_IMAGE`（图路径）与 `DEF_ATT_NEST_DOOR_COUNT`（件序号）
3. 保存总嵌套 ARD（`SaveAllNestARD=False` 时记入 `colDeleteFiles` 事后删除，TFS#56774）
4. 文本转几何（保留 `attSheetIdent`）→ 每板拆单板图纸：**去除引导线**（`PathHasLeadInOut` → `SetLeadInOutNONE/Auto`）、标 `DEF_ATT_POCKET_PATH`（ProcessType 2/4/5 铣槽）、移入临时图纸、保存 `JobName_材料_板名.ard`
5. 报表图：白底 → 每板着色（槽=黄，其余=黑）→ `ZoomToBox` → 隐藏废料（`g_ShowHideScrapCuts`）→ 整板 EMF
6. `EditMark` 高亮区（3867 起）：同件多路径按面积去重 → 当前件红色+`HatchPath` 阴影、其余浅灰 → **逐件保存 EMF**（供报表逐门缩略图）
7. 恢复背景/刷新/undo，重开总图

### mbln_ComparePressNestToOrder（3997–4028）
压机版数量校验：按 `GroupByMaterialThickness` 决定统计范围（单厚度或全厚度），对比嵌套板 `Parts.Count`。

### ⭐ mbln_ProcessPart（4030–4146）— 门型加工分派器
```
Select Case Door.StyleNumber
  900 → mbln_Style_900     （预插入 ARD）
  910 → mbln_Style_910     （ARM）
  920 → mbln_Style_920     （ARB）
  930 → mbln_Style_Make_930（用户自定义宏，bPress 传压机模式）
  Else → 未知门型：预览直接失败；生产询问是否继续
```
失败 → `App.New` 清图 + 询问是否继续（`bStopProcessing` 决定 900 是否直接中止）。

### m_FillDoorData（4148–4217）— 明细→门对象
`DetailID=PK`、`StyleNumber`、`TypeName`、`UserStyleName=StyleName`、宽高/数量/圆角、`RotationMethod/Angle`、`ByPassNest`、`OversizeX/Y`、`NestingZone/ZoneID`、**`CSV_CustomerName/CSV_OrderNumber/CSV_ItemNumber`、`ProductionComment`、`CustomField1/2`、`ComponentGrouping`**、`UserVariableString/UserVariableDescriptionString`、**`UserArg_0~6=UserValue_0~6`**、`HandleID`。预览模式只取几何相关字段（StyleNumber 强制 930）。

### m_FillMaterialData（4219–4270）— 材料行→材料对象
40+ 字段：尺寸/厚度、间隙、`InsertFiller`（填充件文件×3）、`MinimizeToolChanges`、`CutSmallPartsFirst`、`NCSubroutines`、`OnionSkin` 全套、`ProcessWaste` 全套、`PackFinalSheetComponentsToLHS` 等。

### 辅助（4274–4433）
- `m_DeleteReportData`：删订单旧 EMF 图 + `DELETE FROM AD_REPORT_DATA WHERE OrderID=...`
- `m_OutputToAlphaEDIT`：启动 `AlphaEdit.App` 打开所有嵌套/单门 ANC
- `m_UpdateDBNestPaths`：更新报表表嵌套 ANC/ARD 路径（支持按板拆分 `bSplit`）

### mbln_InsertARD（4435–4588）— 900 门型预插 ARD
1. 校验插入文件存在；`g_ConvertGroupsToAttributes` 保存群组号（保存会丢自定义群组）
2. 存当前图为临时文件 → 打开插入文件校验：**工作平面警告**（`WorkPlanes.Count<>0` 报错）、深度容差 `gbln_CheckDepthTolerance`（非预览）、范围取法按选项（`UseDrawingExtentsForInsertedDrawings`，TFS#81512/TFS#59874 大直径刀具问题）、**尺寸适配**（插入件超出门宽/高则失败）
3. 重开原图 → `g_ConvertAttributesToGroups` 恢复群组 → 按参考点插入（`InsertFileReferencePoint` 9 种 + `adoorINSERT_PARAMETRIC` 参数化 `gbln_InsertDrawing`）
4. 新刀路设属性（`m_SetAttributes`）→ 删除多余 work volume


## S7 段分析（4591–5773 行）：单门加工执行 / 更新保存

### 初始化与属性（4591–4660）
- `mbln_Intialize`：`App.New` 新图纸，几何计数归零
- `m_SetAttributes`：给刀路批量标属性——`DEF_ATT_ALPHADOOR`、`DETAIL_ID`、`STYLE_NUMBER`、`TYPE_NAME`、`TOOL_NAME`、`PART_WIDTH/LENGTH/AREA`、`GROUP_ID`、**`CUST_NAME/ORDER_NUM/ITEM_NUM`（来自 CSV_*）**、`ORDER_ID`、`PRODUCTION_COMMENT`、`USER_STYLE_NAME`、`HANDLE_NAME`、`JOB_NAME`、`COMPONENT_GROUPING`；`ByPassNest` 时标板尺寸（宽/长含 Oversize）

### mbln_FillPathData（4661–4797）— 路径配置装载
数据库刀路行 → `clsPathData` 60+ 字段：补偿/弦误差/切削方向、深度体系（FinalDepth、百分比、多刀）、引导线全套（进/出、斜向、圆弧半径、重叠）、加工方法、`ToolFullPath=LicomdatPath & ...`、刀具号（`UseToolData=False` 时从 DB 取号/转速/进给）、插入文件（路径/参考点/群组号）、创建方法（默认 MANUAL）、切削类型（默认 FULL，兼容旧数据空串）、部分切削 4 参数、拐角减速 6 参数、简单雕刻。

### m_SetPostVariables（4799–4822）
后处理变量：`PROGNUM`（4 位零填充）、`DESCRIPTION`（日期）、`ALPHADOOR=1`、`ACAM_CDM=1`、`PANEL_X/Y/Z`、`ZDIM`。

### ⭐ mbln_MakeMachining（4824–5137）— 单门加工执行
按 `grst_GetPaths(Door.TypeName)` 刀路配置逐条执行：
1. **INSERT 方法** → `mbln_InsertARD`（失败询问继续）
2. **常规方法**：取待加工几何——有 `GroupID` → `mpths_PathsInGroup`；无 → `mpths_PathsNotInGroup`（`IgnoreOuterGeometry` 时不加工外框，LJO 08.02.11 修复）
3. 逐路径：`PathOffsetValue≠0` → `mbln_MachineWithOffset`（S8）；否则部分切削 → `DrawParametricPartialPath` 或原几何 → `mbln_MachineWithNoOffset`（S9）
4. 参数应用：`ToolDirectionIsReversed` → `Reverse`；闭几何 `ToolInOut` / 开几何 `ToolSide`；`CW`；选路径；边界（`pthBoundary.CW=Not ToolDirectionCW`）
5. `CreationMethod=MANUAL` → `gbln_GetTool`；`gbln_MakeMillData` 生成刀路 → 粗精/雕刻/槽/简单雕刻加引导线 `gbln_ApplyLeads` → `m_SetAttributes` → 粗精拐角减速 `SlowDownForCorners` → 临时路径删除（`DEF_ATT_TEMP_PATH`）
6. 收尾：`mbln_DrawHandleHoles` 手柄孔 → **非 CHINACAM 级别**（TFS#62973）时运行 `clsOptions.CustomMacro`（`mbln_RunCustomMacro`，TFS#56993）→ `mbln_UpdateAndSave`

### ⭐ mbln_UpdateAndSave（5139–5583）— 保存与入嵌套列表
1. 旋转/纹路处理：`RotationMethod`（FREE 不动 / LOCKX 转 90° 锁 0 / LOCKY 锁 0 / MATERIAL 按 `AllowRotation+GrainInX`）＋ `g_PartRecovery` 部件回收偏移（`PartRecoveryOffsetX/Y`、`IgnoreGrain`）
2. 900 门型旋转限制从图纸属性 `DEF_ATT_NEST_RESTRICT` 读取；930 门型应用规则 `mbln_RulesApply`
3. 文件名 `JobName_TypeName_序号`；双位置保存 ARD/NC/MPR（`ByPassNest` 单件模式**总是输出 NC**，文件名含 `宽x高`）
4. 每个刀路标 `ANC_NAME/ANC_FULLNAME/PART_IMAGE`、`DEF_ATT_DOOR_ROTATION_ANGLE`（供 Twin Head 嵌套）
5. 旋转门 `ByPassNest` → 平移 `MoveL Door.Length, 0` 到正象限
6. 输出 EMF 报表图 → 保存 ARD → **RADNEST 自动分区**（TFS#79915：无手动分区时按 `ZonePartDimension/Area` 自动分配）→ `clsNest.AddPart(路径, 数量, 旋转角, 优先级, 分区)` 加入嵌套列表 → `SaveAllDoorARD=False` 登记删除

### mbln_UpdateAndSavePress（5585–5770）— 压机版
旋转按 `RotationMethodPress` + 颜色级 `ColourRotationMethod/ColourDefinedRotationMethod` 两级；文件名 `JobName_PressName_颜色_厚度_门型_序号`（900 用 `gstr_ParseName`，930 用 TypeName）；ARD 保存 + `clsNest.AddPart(..., 0)`。

### m_SmallFirst（5773 起）
小件先切排序：Rectangular 嵌套不保证小件最后，真空损失导致小件移动（Spacemaker/OS Doors 案例）——`SuppressUpdateRapids True` 后重排外部刀路。


## S8 段分析（5773–6990 行）：小件排序 / 填充件 / 报表写入 / 订单信息 / NC 拆分 / 偏移

### m_SmallFirst 尾部（5781–5786）
`OrderExternalToolpaths` 重排外部刀路（小件先切）。

### mbln_InsertFillerParts（5788–6013）— 边角料填充件
1. 校验 `Material.InsertFillerFile`（1~3 个填充 ARD 文件，`LicomdirPath` 下，去扩展名；全部缺失 → 提示失败）
2. 对每块嵌套板：`StartNestListRouter("材料_FILL")` → `AddPart(填充件, 500, 90, 1, 0)`（数量 500 模拟 MAX）→ `CreateNestData` 配置（方向 TOPRIGHT、EdgeGap/Gap/LeadGap/Resolution/Subroutines/MergeTools/InnerFirst）→ `DoNest` 填充剩余空间
3. ⚠️ 含约 110 行注释掉的旧版方案（5903–6012 死代码）

### m_InsertReportDataPress（6015–6166）— 压机报表写入
遍历嵌套板：统计同名件在板数量 → 给刀路标 `DEF_ATT_PRESS_NAME/SHEET_NAME/ITEM_NUMBER/QTY_ON_SHEET/SHEET_NUMBER` → `INSERT INTO AD_REPORT_DATA`（DetailID/OrderID/CustomerID/PressName/FoilColour/PressSheetName/件号/板内数量/PressPathToEMF/PressSheetNumber/PathToPressNestARD）→ 最后 `m_CreateAlphaCAMDrawingsOfSheetsPress`。
⚠️ **疑似 bug（6106 行）**：`DEF_ATT_PRESS_SHEET_NUMBER` 被赋值两次，第二次覆盖为 `PressColour.ColourName`（应为不同属性名）。

### ⭐ m_InsertReportDataRouter（6168–6456）— 路由器报表写入
1. `m_CreateAlphaCAMDrawingsOfSheets` 生成单板图 → 遍历嵌套板：板条码 `gstr_GenerateNestedSheetBarcode`（标 `SHEET_UNIQUE_ID`）、废料率 `gdbl_Scrap`
2. 遍历零件刀路：AlphaDOOR 属性（`DEF_ATT_ALPHADOOR=1`）→ 取 `ANC_NAME/ANC_FULLNAME/PART_IMAGE/ITEM/NEST_DOOR_IMAGE/COUNT`、**`CUST_NAME/ORDER_NUM/ITEM_NUM`（CSV 客户/订单/件号）**、`PRODUCTION_COMMENT`、`FoilColour`；非 AlphaDOOR → 视为填充件（`DEF_SCRAP_FILLER`）
3. `AdditionalReportData` → 取零件几何范围（`NestPartPositionLeft/Right/Top/Bottom`，`gs_NoComma` 防 SQL 逗号问题）
4. **压机联动**：`colReportData` 命中 → `GetNextID` + `UpdatePressInfo`（回填压机数量，S4 已述）；否则 INSERT（30+ 字段含 Sheet 全套/Part 全套/ANC/ARD/EMF/条码/压机图/位置）；有 PK → UPDATE 同字段

### m_FillOrderInfo（6461–6513）
订单 + 客户信息 → 门对象：`OrderID/JobName/ProcessedDate/CustomerID/HotJob/DueDate/OrderDate/PO` + 客户地址/联系人/电话/邮箱（无客户记录置空串）。

### ⭐ m_SplitNestANC（6515–6749）— 多板 NC 拆分
按板把整板 NC 拆成单板程序：
1. 板名集合（`DetermineSheetID`）；删旧 `Worklist.seq` / `Pallet.txt`
2. 扫描 NC 找 `START_FLAG`（`mbln_HasSTART` 判定原文件是否以 START 开头）→ 遇 START 开新文件 `原文件名_板号.nc`（重插 START + `SplitProgramsStartOfFileMarker`）→ 直至 `END_FLAG` 截断（补 `EndOfFileMarker`）
3. 可选 `SplitProgramsOutputPGM` → `m_ConvertToPGM` + `m_CheckForErrors`
4. 每板 `m_UpdateDBNestPaths(..., bSplit=True, 板号)`；`SplitControlFile` → 追加 Worklist.seq；`SplitAutoLoad` → 按材料匹配托盘 1/2 写 `Pallet.txt`（文件名 79 字符补齐）；收集 `colANC`

### 辅助（6751–6962）
- `mbln_HasSTART`：NC 首 5 字符是否 "START"
- `m_ConvertToPGM`：调 `winxiso.exe <ShortPath> -s -i`（`ShellAndWait`）
- `m_CheckForErrors`：读 winxiso 输出 `.INF` 末行，`=0` 无错（可选删除 INF），否则弹错
- `mpths_PathsInGroup`：按 `Group` 筛路径
- `mpths_PathsNotInGroup`：按 `DEF_ATT_GEOMETRY_NUMBER = PathOffsetFrom` 取首个匹配

### mbln_MachineWithOffset（6964–6990 起）
偏移加工入口：闭路径按 CW 与 `PathOffsetSide`（内/外）组合 `Offset(值, acamLEFT/RIGHT)`；非部分切削时 `gbln_SetStartPoint` 设起点。


## S9 段分析（6990–7926 行）：偏移加工 / 属性 / 门型 Style 900–930

### 偏移加工（6964–7250）
- `mbln_MachineWithOffset`：闭路径按 `CW` × `PathOffsetSide`（内/外）四种组合 `Offset(值, acamLEFT/RIGHT)`；开路径固定内 RIGHT / 外 LEFT。部分切削 → `DrawParametricPartialPath`（删原几何）；`PocketBoundary≠0` → 内偏 `PocketBoundary - ToolDiameter/2` 生成深红边界；偏移结果棕色
- `mbln_MachineWithNoOffset`：设起点 + 按 `pthPick.CW` 方向内偏边界

### 属性批量设置（7252–7317）
- `m_SetDetailAttributes`：`LicomUSrlg_alphadoor_*` 用户属性 25+ 项（客户/地址/PO/交期/热单/门型/宽高/圆角/UserStyleName/UserVariableString/**UserArg_0~6**/ProductionComment/Custom1/2/FoilColour）
- `m_SetPressAttributes`：压机精简属性（ALPHADOOR/DETAIL_ID/ORDER_ID/CUSTOMER_ID/TYPE_NAME/FOIL_COLOUR）

### 门型加工族（7320–7899）

| 门型 | 过程 | 机制 | 差异点 |
|------|------|------|--------|
| 900 | `mbln_Style_900` | 预插入 ARD 图纸（`App.OpenDrawing .TypeName`） | 实际宽高从图纸范围 `GetExtentL` 反算；有手柄孔 + 自定义宏；压机模式删刀路只存几何 |
| 910 | `mbln_Style_910` | `App.RunParametricMacro(.TypeName)` 参数化宏 | 无手柄孔/自定义宏步骤 |
| 920 | `mbln_Style_920` | `App.Run(.TypeName & ".ADOOR.MakeDoor")` 工程宏 | 先 `gbln_ProjectExists` 校验 |
| 930 | ⭐ `mbln_Style_Make_930` | 用户 VBA 工程宏（`UserStyleName`） | 见下 |

**930 用户门型核心流程**：
1. `gbln_ProjectExists(.UserStyleName)` 校验宏工程（README 所述"无法连接用户定义的宏"报错即源于此，`-2147467259` → 帮助）
2. 组装 `CUserStyle`：宽/长/圆角 + `UserVariableString` 按 `";"` 拆分入 `UserVariables` + 描述拆分
3. 非忽略外框时：`CreateRectangle(0,0,宽,长)` 画外围 → **`StdAlpha.ShareClass` DLL 扩展** `dll.jc`（用户变量 46–49 控制倒角/切角）→ `dll.diamond`（圆角）→ 几何号 1 → `m_SetDetailAttributes`
4. `App.Run(.UserStyleName & ".ADOOR.MakeDoor", RequiredData, UserArg_0..6)` 运行宏 → 检查 `PathsToReturn` 与 `Success`
5. 宏返回几何逐个标几何号 + 详细属性（第一个压机几何补 Press 属性，防宏删外框）
6. 清 work volume → 非压机 `mbln_MakeMachining`（含 S7 全部加工链）/ 压机 `mbln_UpdateAndSavePress`
7. 收尾 `UserVariables(50)≠0` → `dll.bz`（DLL 附加处理）

### 测试桩（7901–7926）
`Temp`（遍历 Operations 测试）、`TestScrap`（`Debug.Print` 各板废料率）——遗留调试代码，无生产价值。

## 三、跨段调用链（g_Make_Master 全链路）

```
g_Make_Master(strOrder)  [S4]
 ├─ m_PopulatePressData / m_PopulateRouterData   [S5] ← modAutoImportNest 导入数据入口
 ├─ gbln_Make_Master_Press → m_DoNestingPress    [S2/S3] 压机支线
 ├─ For Each Router/Material/Door
 │    ├─ mbln_ProcessPart → Style 900/910/920/930 [S6/S9]
 │    │    ├─ mbln_Style_Make_930 → mbln_MakeMachining [S7]
 │    │    │    ├─ mbln_FillPathData → mbln_MachineWith(No)Offset [S8/S9]
 │    │    │    ├─ mbln_DrawHandleHoles [S1]
 │    │    │    └─ mbln_RunCustomMacro / mbln_UpdateAndSave [S1/S7]
 │    │    └─ mbln_UpdateAndSave → clsNest.AddPart [S7]
 │    └─ m_DoNestingRouter [S3]
 │         ├─ N.Nest / 校验 CompareRouterNestToOrder [S3]
 │         ├─ m_PackFinalNestSheetToLHS / InsertFillerParts / SmallFirst [S1/S8]
 │         ├─ 输出 NC/MPR/NCHops + m_SplitNestANC + PGM 转换 [S3/S8]
 │         └─ m_InsertReportDataRouter → m_CreateAlphaCAMDrawingsOfSheets [S8/S6]
 └─ 报表回填 / COMPLETENEST 标志 / 恢复后处理  [S4]
```

## 四、发现的问题与风险

1. **6106 行属性覆盖 bug**：`m_InsertReportDataPress` 中 `DEF_ATT_PRESS_SHEET_NUMBER` 连续赋值两次，第二次误写 `PressColour.ColourName`（疑似应为别的属性名），压机板号属性丢失
2. **死代码多**：S2 中约 200 行注释旧嵌套逻辑（1404–1603）、S8 约 110 行旧填充件方案（5903–6012）、`mbln_Intialize` 几乎空实现、`Temp`/`TestScrap` 调试桩——可清理但无功能影响
3. **第三方 DLL 依赖**：`StdAlpha.ShareClass`（`jc`/`diamond`/`bz`/`ScrapNesting`）——若 DLL 缺失或接口变化，930 外框/边角料/填充件功能直接失败，且无降级路径
4. **外部工程耦合**：`NcHopsPost`、`WoodWop4xV20_edge/normal`、`AlphaDOOR_ReverseSideNesting` 通过 `App.Run` 字符串调用，工程名不匹配时静默失败（`m_NCHopsOutput` 有检查，`m_MprSave` 无）
5. **SQL 拼接**：报表 INSERT/UPDATE 全字符串拼接（`gs_FixSQL`/`gs_NoComma` 防护有限），特殊字符字段有注入/截断风险
6. **旋转/纹路逻辑复杂**：`mbln_UpdateAndSave` 中 LOCKX/MATERIAL/`g_PartRecovery`/`blnRotateDoor` 多重分支，900/930 行为不一致（930 由宏负责旋转）

## 五、与 modAutoImportNest / frmAutoNest 的关联

- **数据入口**：导入的 `AD_ORDER_DETAILS` → `m_PopulateRouterData`（S5）→ `m_FillDoorData` 读取 `StyleNumber/UserStyleName(UserStyleName)/UserValue_0~6/CSV_*/ComponentGrouping` 等 → 门型分派（S6）→ 加工/嵌套/报表
- **930 门型成败**：`mbln_Style_Make_930` 依赖 `UserStyleName` 匹配宏工程名（`gbln_ProjectExists`）与 `UserVariableString`/`UserValue_0~6`（宏参数）——正是 README/导入模块反复强调的"930 复制 UserStyleName + UserValue_0~6"的原因
- **成败标志**：`COMPLETENEST` 注册表（-1/0）被 `AutoImportNest`/`AutoImportNestWithParams` 读取用于完成提示
- **NC/ARD/EMF 产物**：`m_DoNestingRouter` 与 `m_SplitNestANC` 生成的程序、单板图、报表图构成全部交付物，路径回写 `AD_REPORT_DATA` 供报表
