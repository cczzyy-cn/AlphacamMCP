# CDM 模块源码分析报告

> 基于 AI工作流程/ 目录中的完整源码分析生成
> CDM.arb 版本 3.0 | 数据库版本 1.3 | CDM2016R1

---

## 一、CDM 整体架构

```
AlphaCAM 2016 R1
    │
    ├── CDM.DLL (4.6MB COM 组件) ─── 主界面 + 右键菜单
    │   ├── MainFrontEnd.RunProcessing(App) ─ 三面板 UI
    │   ├── Import CSV (导入) ─ 内部处理，无 VBA 源码
    │   ├── Run Production (批量生产) ─ 调用 Make.g_Make_Master
    │   └── 门库树/订单树/生产队列
    │
    ├── CDM.arb (VBA 插件) ─── 加工逻辑 + 数据库
    │   ├── Events.bas         → 菜单注册 + CDM 入口
    │   ├── Make.bas (7926行)  → 核心加工引擎
    │   ├── frmNTCW (3955行)   → 门型创建向导
    │   ├── CNest (1023行)     → 排版列表生成
    │   ├── ProcessWaste       → 废料处理逻辑
    │   ├── Database           → DB 操作层(受保护)
    │   ├── Functions          → 工具函数(受保护)
    │   ├── Globals            → 全局变量/枚举
    │   └── 30+ 类模块         → CDoor/CMaterial/CRouter...
    │
    ├── CDM.mdb ─── Access 数据库 (20 个表)
    └── CDM.ctx ─── 界面文字资源
```

---

## 二、模块详细分析

### 2.1 Events.bas (2852 行/108KB)

#### Public 函数

| 函数 | 行号 | 功能 | 可被 run_vba_macro 调用？ |
|------|------|------|--------------------------|
| `InitAlphacamAddIn(AcamVersion)` | 175 | 注册 CDM 菜单到 AlphaCAM | ✅ AlphaCAM 启动时自动调用 |
| **`m_Processing()`** | **2137** | **打开 CDM 主界面（CDM.DLL 接管）** | ✅ `"CDM.m_Processing"` |
| `m_OrderToolPaths()` | 113 | 排序刀具路径 | ✅ `"CDM.m_OrderToolPaths"` |
| `m_TestUserStyles()` | 372 | 测试用户样式 | ✅ `"CDM.m_TestUserStyles"` |
| `m_About()` | 2254 | 关于对话框 | ✅ |
| `m_Help()` | 2262 | 打开帮助 | ✅ |
| `g_ConfigDB()` | 2822 | 数据库连接配置 (UDL) | ✅ |
| `g_CompactDB()` | 310 | 压缩数据库 | ✅ |
| `g_SupportUtil()` | 345 | 技术支持工具 | ✅ |
| `mbln_BackupDatabase()` | 133 | 数据库备份 | ✅ |

#### 关键代码

```vb
' m_Processing — CDM 主入口 (第2137行)
Function m_Processing()
    ' 创建 CDM.DLL COM 对象
    mbln_CreateObject oCDM, "CDM2016R1.MainFrontEnd", "CDM.DLL"
    ' 运行 CDM 主界面
    oCDM.RunProcessing App
End Function

' InitAlphacamAddIn — 注册菜单 (第175行)
Function InitAlphacamAddIn(AcamVersion As Long) As Integer
    ' 注册 10 个菜单项：
    ' m_Processing(主界面), m_TestUserStyles, m_OrderToolPaths,
    ' g_ConfigDB, g_CompactDB, g_SupportUtil, m_Help, m_About
End Function
```

---

### 2.2 Make.bas (7926 行/290KB) ⭐ 核心

#### Public 入口函数

| 函数 | 行号 | 功能 |
|------|------|------|
| **`g_Make_Master(strOrder)`** | **2274** | **🎯 批量生产总入口：加工→排版→NC** |
| `gbln_Make_Master_Press()` | 1307 | Press 压机加工入口 |
| `m_PopulatePressData(OrderString)` | 2811 | 从 DB 读取 Press 数据 |
| `m_PopulateRouterData(OrderString)` | 2925 | 从 DB 读取 Router 数据 |
| `mbln_PopulateReportData(OrderString)` | 3004 | 生成报表数据 |
| `g_Preview_Master(lType)` | 3267 | 预览门型 |
| `m_CreateAlphaCAMDrawingsOfSheets(Material)` | 3632 | 创建排版后图纸 |
| `m_NCHopsOutput(...)` | 373 | NC Hops 输出 |
| `mbln_CompareRouterNestToOrder(...)` | 2162 | 比较排版与订单 |
| `mstr_GetJobName(OrderID)` | 3250 | 获取订单名 |
| `Temp()` | 7901 | 测试函数 |
| `TestScrap()` | 7916 | 废料测试 |

#### Private 核心函数链

```vb
' g_Make_Master 内部流程 (第2274行)
Public Sub g_Make_Master(strOrder As String)
    ' 1. 初始化：连接 DB、清空旧报表数据
    ' 2. 对每个 Router（加工配置）：
    '    对每个 Material（板材）：
    '       a) App.New → 新建图纸
    '       b) clsNest.StartNestListRouter → 创建排版列表头
    '       c) 对每个 Door（门板）：
    '          mbln_ProcessPart(Door, ...)  ← 核心加工
    '          clsNest.AddPart              ← 加入排版
    '       d) m_DoNestingRouter → 执行排版
    '       e) NC 输出
    ' 3. 完成报告
End Sub

' mbln_ProcessPart — 单个门板加工 (第4030行)
Private Function mbln_ProcessPart(Door, PartNumber, colANC, bPress, bPreview)
    Select Case Door.StyleNumber
        Case 900: mbln_Style_900     ' 标准镶板门 — 画矩形+刀路
        Case 910: mbln_Style_910     ' 插入 ARM 宏文件
        Case 920: mbln_Style_920     ' VBA 宏
        Case 930: mbln_Style_Make_930 ' 用户自定义
        Case Else: MsgBox "未知门型!" ' ❌ StyleNumber 错误会弹窗
    End Select
End Function
```

---

### 2.3 frmNTCW (3955 行/124KB)

**门型创建向导** — 不是 CDM 主界面！

Wizard 步骤：
```
0. DOOR_TYPE → 门型名称
1. CREATION_METHOD → 手动/加工样式
2. MACHINE_METHOD → RoughFinish/Pocket/Engrave/Insert
3. INSERT_PATH → 插入文件参数
4. OFFSET_AMOUNT → 偏移量
5. TOOL_DIRECTION_SIDE → 刀具方向/侧
6. POCKETING → 型腔参数
7. ROUGH_FINISH → 粗精加工
8-10. MACHINING_1/2/3 → 加工参数
11. LEADS → 引入引出
12. POCKET_LEADS → 型腔引入引出
13. MACHINING_STYLES → 加工样式
14. PATH_COMPLETE → 完成
```

---

### 2.4 CNest (1023 行)

#### Public 函数

| 函数 | 功能 |
|------|------|
| `StartNestListRouter(JobName, Material, FillerPart, PackLHS)` | Router 排版总入口 |
| `StartNestListPress(JobName, Press)` | Press 排版入口 |
| `AddPart(PartName, Qty, Rotation, Priority, NestZone)` | 添加零件到排版列表 |
| `NestingOption` (Property) | 排版方式 (TrueShape/RadNest/Rect/Disabled) |

排版方式选择：
```vb
Sub StartNestListRouter(...)
    Select Case NestingOption
        TrueShape: StartNestListTrueRouter  ' 真形排版
        RadNest:   StartNestListRadnestRouter ' RadNest
        Rect:      StartNestListRectRouter    ' 矩形排版
        Disabled:  不排版
    End Select
End Sub
```

---

### 2.5 ProcessWaste (2568 行)

废料处理 — 铣去排版后零件之间的废料区域。

| 函数 | 功能 |
|------|------|
| `ProcessWasteMaterial(Style, Depth, FinalScrap, CutTowards, Strategy, HSpacing, VSpacing)` | 🎯 主入口 |
| `g_ProcessWasteRoughFinish(...)` | 粗精加工方式废料处理 |
| `OptimiseNestedScrapCuts()` | 优化废料刀路排序 |
| `g_ShowHideScrapCuts(Show)` | 显示/隐藏废料刀路 |

两种策略：
```vb
AdoorProcessWasteStrategy_Pocketing = 0   ' 型腔铣削
AdoorProcessWasteStrategy_RoughFinish = 1 ' 粗精加工
```

---

## 三、数据库 (CDM.mdb) — 20 个表

### 3.1 核心表结构

#### AD_ORDERS — 订单
```
OrderID (Long, PK)    JobName (VarChar)     CustomerID (Long)
OrderDate (VarChar)   DueDate (VarChar)     ProcessedDate (VarChar)
```

#### AD_ORDER_DETAILS — 门板明细 ⭐
```
PK (Long, PK)         OrderID (Long)        StyleNumber (Long)
TypeName (VarChar)    Quantity (Long)        Width (Double)
Length (Double)       Material (VarChar)    CornerRadius (Double)
ProductionComment     PressID, ColourID     CSV_CustomerName
CSV_OrderNumber       CSV_ItemNumber        HandleID
...共 49 个字段
```

#### AD_DOOR_TYPES — 门型定义
```
PK (Long, PK)         TypeID (VarChar)      StyleNumber (Long) ← 关键!
Width, Length         PressID, ColourID     HandleID
...共 29 个字段
```

#### AD_DOOR_PATHS — 刀路参数 ⭐
```
PathID, TypeID,       PathNumber (工序号)   ToolNumber
MachiningMethod       SafeRapid, RapidDownTo, FinalDepth
McComp, SpindleSpeed, DownFeed, CutFeed
LeadIn, LeadOut,      PocketType, WidthOfCut
...共 79 个字段
```

#### AD_MATERIALS — 材料
```
Name, Width, Length, Thickness            NoRotation, PackTo
SearchRes, MinGapBetweenPaths, LeadGap   OnionSkin 系列
ProcessWaste 系列                         InsertFiller
...共 41 个字段
```

### 3.2 当前数据库状态

| 表 | 状态 |
|----|------|
| AD_MATERIALS | ✅ 已有 14 种材料 |
| AD_DOOR_TYPES | ❌ 空 (需要 BatchImport 自动创建) |
| AD_DOOR_PATHS | ❌ 空 (需在 CDM UI 中预配) |
| AD_ORDERS | ❌ 空 (通过 BatchImport 写入) |
| AD_ORDER_DETAILS | ❌ 空 (通过 BatchImport 写入) |

---

## 四、与自定义模块的关系

### 4.1 BatchImport.bas (已完成)

```vb
Public Sub Run(CSVPath, JobName)
    ' 1. gbln_ConnectToDB → 连接数据库
    ' 2. 解析 CSV 文件 (14列)
    ' 3. glng_GetOrCreateCustomer → 获取/创建客户
    ' 4. glng_CreateOrder → INSERT INTO AD_ORDERS
    ' 5. 逐行: InsertOrderDetail → INSERT INTO AD_ORDER_DETAILS
    '    ├── glng_GetOrCreateStyle → INSERT INTO AD_DOOR_TYPES
    '    └── glng_EnsureMaterial → INSERT INTO AD_MATERIALS
End Sub
```

⚠️ **需修复**：`glng_GetOrCreateStyle` 中 `StyleNumber` 应为 `900`（不是 `1`）

### 4.2 BatchProcess.bas (已完成)

```vb
Public Sub RunByName(JobName)
    ' 1. 查找 OrderID
    ' 2. 读取 AD_ORDER_DETAILS 门板数据
    ' 3. 对每个门板: App.New + CreateRectangle + SaveAs .amd
    ' 注: 跳过 CDM 加工引擎，仅生成几何
End Sub
```

### 4.3 推荐完整工作流

```python
# 1. 安装模块（首次）
install_vba_module("BatchImport", open("CDM功能/BatchImport.bas").read())
install_vba_module("BatchProcess", open("CDM功能/BatchProcess.bas").read())

# 2. 导入 CSV → 创建订单
run_vba_macro("CDM.BatchImport.Run", [...csv..., ...jobname...])

# 3. 直接调用 CDM 加工引擎（推荐，替代 BatchProcess）
run_vba_line("g_Make_Master(" & orderID & ")")
```

---

## 五、已知限制

| 限制 | 说明 |
|------|------|
| CDM.DLL 不可读 | 主界面 UI + CSV 导入 + 批量生产调度无法直接获取源码 |
| Database 模块受保护 | `gbln_ConnectToDB` 等核心函数无法读取 |
| Functions 模块受保护 | `gs_FixSQL`、`gvar_CheckNull` 等工具函数无法读取 |
| frmNTCW 非主界面 | 只是门型创建向导，不是三面板主界面 |
| StyleNumber=1 错误 | BatchImport 应设 900 而非 1 |
