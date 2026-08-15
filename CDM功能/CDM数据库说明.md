# CDM 数据库说明（2026-08-15）

> 数据来源：AlphaCAM `gdb_CDM` 实查 `CDM.mdb`（OpenSchema 导出 20 张表）；字段用途对照 `CDM功能/Make.bas`、`modAutoImportNest.bas`、`Events.bas` 代码标注。
> 数据库文件：`CDM.mdb`（约 1.2GB，Access/Jet 或 SQL Server，由 `gbln_SQLServer` 区分）

## 一、表索引（20 张）

| 表 | 用途 |
|----|------|
| `AD_CUSTOMERS` | 客户 |
| `AD_DOOR_PATHS` | 刀路参数（按门型配置） |
| `AD_DOOR_TREE` | 门型树 |
| `AD_DOOR_TYPES` | 门型定义（含用户样式宏名） |
| `AD_DOOR_XKSET` | 门型附加设置 |
| `AD_HANDLE_DRILLING` | 手柄钻孔配置（mbln_DrawHandleHoles 读取） |
| `AD_MATERIAL_COLOURS` | 材料颜色（膜）配置 |
| `AD_MATERIALS` | 材料定义 |
| `AD_NEST_ZONES` | 排版区域（自动/手动） |
| `AD_OPERATORS` | 操作员 |
| `AD_ORDER_DETAILS` | 门板明细（modAutoImportNest ImportCSV 写入） |
| `AD_ORDER_GRID` | 订单表格配置 |
| `AD_ORDERS` | 订单主表（modAutoImportNest 写入） |
| `AD_PRESSES` | 压机配置 |
| `AD_REPORT_DATA` | 报表/标签数据（Make 排版后写入，标签打印数据源） |
| `AD_RULES` | 路由器规则（m_PopulateRulesCollection 读取） |
| `AD_SETTINGS` | 系统设置 |
| `AD_TOOL_ORDER` | 刀具排序顺序（mbln_ToolOrdering 读取） |
| `AD_USER_STYLES` | 用户样式宏工程（.arb 路径） |
| `AD_VERSION` | 版本信息 |

## 二、核心表字段详解

### AD_ORDERS

| 字段 | 类型 | 长度 | 说明 |
|------|------|------|------|
| CustomerID  | Integer |  | 客户ID（→AD_CUSTOMERS） |
| DueDate  | WChar | 50 | 交期 |
| HotJob  | Boolean | 2 |  |
| JobName  | WChar | 255 | 订单名（唯一，重名取消/覆盖） |
| OrderDate  | WChar | 50 | 下单日期 |
| OrderID 🔑 | Integer |  | 主键 |
| PO  | WChar | 255 |  |
| ProcessedDate  | WChar | 50 | 处理日期（g_Make_Master 完成后写入） |

### AD_ORDER_DETAILS

| 字段 | 类型 | 长度 | 说明 |
|------|------|------|------|
| ByPassNest  | Boolean | 2 | 跳过排版 |
| ColourID  | Integer |  |  |
| ColourRotationMethod  | Integer |  |  |
| ComponentGrouping  | Integer |  | 分组编号（Long，CSV颜色Val转换） |
| CornerRadius  | Double |  | 圆角 |
| CSV_CustomerName  | WChar | 100 | CSV客户名 |
| CSV_ItemNumber  | WChar | 50 | CSV件号 |
| CSV_OrderNumber  | WChar | 100 | CSV订单号 |
| CustomField1  | WChar | 255 | 自定义1（开启方向） |
| CustomField2  | WChar | 255 | 自定义2（终端地址） |
| HandleID  | Integer |  |  |
| IgnoreOuterGeometry  | Boolean | 2 | 忽略外框 |
| Length  | Double |  | 高/长(mm) |
| Material  | WChar | 255 | 材料名（→AD_MATERIALS） |
| NestingPriority  | Integer |  |  |
| NestZoneID  | Integer |  |  |
| OrderID  | Integer |  | 订单ID |
| OriginalByPassNest  | Boolean | 2 |  |
| OversizeX  | Double |  |  |
| OversizeY  | Double |  |  |
| PK 🔑 | Integer |  | 主键 |
| PostProcessor  | WChar | 255 |  |
| PressID  | Integer |  |  |
| Processed  | Boolean | 2 |  |
| ProductionComment  | WChar | 255 | 备注 |
| Quantity  | Integer |  | 数量 |
| ReverseMachiningFilename  | WChar | 255 |  |
| ReworkQuantity  | Integer |  |  |
| RotationAngle  | Double |  | 旋转角 |
| RotationMethod  | SmallInt |  | 旋转方式 |
| SFDC_Approve  | Integer |  |  |
| SFDC_Approve_Stage_2  | Integer |  |  |
| SFDC_Approve_Stage_3  | Integer |  |  |
| SFDC_Reject  | Integer |  |  |
| StyleName  | WChar | 255 | 用户样式名（=AD_DOOR_TYPES.UserStyleName，930门型宏项目名） |
| StyleNumber  | Integer |  | 门型编号（900/910/920/930） |
| TypeName  | WChar | 255 | 门型类型名（→AD_DOOR_TYPES.TypeID） |
| UserDescriptionString  | WChar | 0 | 用户变量描述串 |
| UserValue_0  | WChar | 255 |  |
| UserValue_1  | WChar | 255 |  |
| UserValue_2  | WChar | 255 |  |
| UserValue_3  | WChar | 255 |  |
| UserValue_4  | WChar | 255 |  |
| UserValue_5  | WChar | 255 |  |
| UserValue_6  | WChar | 255 |  |
| UserVariableString  | WChar | 0 | 用户变量串(分号分隔，宏参数) |
| Valid  | Boolean | 2 |  |
| Width  | Double |  | 宽(mm) |

> 注：`UserValue_0`~`UserValue_6` 为宏参数（`UserArg_0`~`UserArg_6`），930 用户门型由 `App.Run` 传给宏。

### AD_DOOR_TYPES

| 字段 | 类型 | 长度 | 说明 |
|------|------|------|------|
| ByPassNest  | Boolean | 2 |  |
| ColourID  | Integer |  | 颜色ID |
| ColourRotationMethod  | Integer |  |  |
| Comment  | WChar | 0 |  |
| CornerRadius  | Double |  |  |
| CreatorName  | WChar | 255 |  |
| DateAdded  | Date |  |  |
| HandleID  | Integer |  | 手柄配置ID |
| IgnoreOuterGeometry  | Boolean | 2 |  |
| Length  | Double |  | 默认长 |
| OversizeX  | Double |  |  |
| OversizeY  | Double |  |  |
| PK  | Integer |  |  |
| PressID  | Integer |  | 压机ID |
| RotationAngle  | Double |  |  |
| RotationMethod  | SmallInt |  |  |
| TypeID 🔑 | WChar | 255 | 门型标识（主键，如 PETA系列） |
| UserDescriptionString  | WChar | 0 |  |
| UserStyle  | Boolean | 2 | 是否用户样式 |
| UserStyleName  | WChar | 255 | 用户样式宏项目名 |
| UserValue_0  | WChar | 255 |  |
| UserValue_1  | WChar | 255 |  |
| UserValue_2  | WChar | 255 |  |
| UserValue_3  | WChar | 255 |  |
| UserValue_4  | WChar | 255 |  |
| UserValue_5  | WChar | 255 |  |
| UserValue_6  | WChar | 255 |  |
| UserVariableString  | WChar | 0 |  |
| Width  | Double |  | 默认宽 |

### AD_DOOR_PATHS

| 字段 | 类型 | 长度 | 说明 |
|------|------|------|------|
| AccelerateOutOfCorner  | Boolean | 2 |  |
| ChordError  | Double |  |  |
| CompOnRapid  | Boolean | 2 |  |
| CreationMethod  | WChar | 50 |  |
| CutDirection  | Double |  |  |
| CutFeed  | Double |  | 切削进给 |
| CutType  | WChar | 50 |  |
| DecelerationDistance  | Double |  |  |
| DepthsOfCutSpecified  | Boolean | 2 |  |
| Diameter  | Double |  |  |
| DoNotSlowDownRadius  | Double |  |  |
| DownFeed  | Double |  | 下刀进给 |
| EngraveCornerAngle  | Double |  |  |
| FinalDepth  | Double |  | 最终深度 |
| FinalDepthPercentage  | Double |  |  |
| FinalPassIsIsland  | SmallInt |  |  |
| GroupID  | SmallInt |  | 几何分组 |
| IgnoreAngleGreaterThan  | Double |  |  |
| InsertFilePath  | WChar | 255 | 插入ARD文件 |
| InsertFilePointX  | Double |  |  |
| InsertFilePointY  | Double |  |  |
| InsertFileReferencePoint  | SmallInt |  | 插入参考点 |
| InsertParametricGroupNumber  | Integer |  |  |
| IsFinalDepthPercent  | Boolean | 2 |  |
| LastModified  | WChar | 50 |  |
| LeadApproachAngle  | Double |  |  |
| LeadArcRadius  | Double |  |  |
| LeadEntryPointIsCorner  | Integer |  |  |
| LeadIn  | SmallInt |  |  |
| LeadLineLength  | Double |  |  |
| LeadLineLengthOut  | Double |  |  |
| LeadOut  | SmallInt |  |  |
| LeadOverlap  | Double |  |  |
| MachiningMethod  | WChar | 50 | 加工方法(INSERT/ROUGHFINISH等) |
| MachiningStyle  | WChar | 255 |  |
| MaterialTop  | Double |  |  |
| McComp  | SmallInt |  |  |
| MultiplePasses  | Boolean | 2 |  |
| NumberOfCuts  | SmallInt |  |  |
| NumberOfSteps  | Integer |  |  |
| PartialEndElemDist  | Double |  |  |
| PartialEndElemIndex  | Integer |  |  |
| PartialStartElemDist  | Double |  |  |
| PartialStartElemIndex  | Integer |  |  |
| PathID 🔑 | Integer |  |  |
| PathNumber  | Integer |  | 工序号 |
| PathOffsetFrom  | Integer |  |  |
| PathOffsetSide  | Integer |  |  |
| PathOffsetValue  | Double |  |  |
| Pocket3DApproach  | Boolean | 2 |  |
| PocketBoundary  | Double |  |  |
| PocketType  | SmallInt |  |  |
| RapidDownTo  | Double |  | 快速下刀 |
| SafeRapid  | Double |  | 安全快速 |
| SimpleEngraveClearance  | Double |  |  |
| SimpleEngraveFeed  | Double |  |  |
| SlopeIn  | Boolean | 2 |  |
| SlopeOut  | Boolean | 2 |  |
| SlowDownForCorners  | Boolean | 2 |  |
| SlowDownTo  | Double |  |  |
| SpindleSpeed  | Double |  | 转速 |
| StartCutting  | SmallInt |  |  |
| StepLength  | Double |  |  |
| Stock  | Double |  |  |
| ThicknessFirstCut  | Double |  |  |
| ThicknessFirstCutPercent  | Double |  |  |
| ThicknessLastCut  | Double |  |  |
| ThicknessLastCutPercent  | Double |  |  |
| ToolDirectionCW  | Boolean | 2 |  |
| ToolDirectoinReversed  | Boolean | 2 |  |
| ToolFullPath  | WChar | 255 |  |
| ToolInOut  | SmallInt |  |  |
| ToolName  | WChar | 255 | 刀具名 |
| ToolNumber  | Integer |  | 刀具号 |
| ToolOffset  | Integer |  | 刀补号 |
| ToolSide  | SmallInt |  |  |
| ToolSidePartialReverse  | Boolean | 2 |  |
| TypeID  | WChar | 255 | 门型ID |
| WidthOfCut  | Double |  |  |
| XYCorners  | SmallInt |  |  |

### AD_MATERIALS

| 字段 | 类型 | 长度 | 说明 |
|------|------|------|------|
| CutInnerPathsFirst  | Boolean | 2 |  |
| CutSmallPartsFirst  | Boolean | 2 |  |
| FinalZTolerance  | Double |  |  |
| HorizontalCutSpacing  | Double |  |  |
| InsertFiller  | Boolean | 2 | 是否填充件 |
| InsertFillerFile  | WChar | 255 | 填充件文件 |
| InsertFillerFile2  | WChar | 255 |  |
| InsertFillerFile3  | WChar | 255 |  |
| LeadGap  | Double |  |  |
| LeaveEdgeGapUncut  | Boolean | 2 |  |
| Length  | Double |  | 板长 |
| MaterialDefault  | Boolean | 2 |  |
| MinGapAtSheepEdge  | Double |  | 板边距 |
| MinGapBetweenPaths  | Double |  | 件距 |
| MinimizeToolChanges  | Boolean | 2 |  |
| Name 🔑 | WChar | 255 | 材料名 |
| NCSubroutines  | Boolean | 2 |  |
| NestingScreenUpdate  | Boolean | 2 |  |
| NoRotation  | Boolean | 2 | 禁止旋转 |
| NumberComponentsBySize  | Boolean | 2 |  |
| OnionSkin  | Boolean | 2 | 洋葱皮 |
| OnionSkinApplyToInside  | Boolean | 2 |  |
| OnionSkinCutOrder  | Integer |  |  |
| OnionSkinMinArea  | Double |  |  |
| OnionSkinMinXY  | Double |  |  |
| OnionSkinThickness  | Double |  |  |
| PackFinalSheetComponentsToLHS  | Boolean | 2 | 末板打包LHS |
| PackTo  | SmallInt |  | 打包方向 |
| ProcessWaste  | Boolean | 2 | 废料处理 |
| ProcessWasteCutTowardsComponents  | Boolean | 2 |  |
| ProcessWasteDepthOfCut  | Integer |  |  |
| ProcessWasteFinalSheetScrap  | Double |  |  |
| ProcessWasteMCStyle  | WChar | 255 |  |
| ProcessWasteStrategy  | Integer |  |  |
| RotationFlip  | Boolean | 2 |  |
| SearchRes  | Double |  |  |
| SuppressFinalSort  | Boolean | 2 |  |
| Thickness  | Double |  | 厚度 |
| TimePerSheet  | Integer |  |  |
| ToolPathsOnly  | Boolean | 2 |  |
| VerticalCutSpacing  | Double |  |  |
| Width  | Double |  | 板宽 |

### AD_CUSTOMERS

| 字段 | 类型 | 长度 | 说明 |
|------|------|------|------|
| Address_1  | WChar | 255 |  |
| Address_2  | WChar | 255 |  |
| City  | WChar | 50 |  |
| Contact  | WChar | 255 |  |
| CustomerID 🔑 | Integer |  | 主键 |
| Email  | WChar | 255 |  |
| Fax  | WChar | 50 |  |
| Name  | WChar | 255 | 客户名 |
| State  | WChar | 50 |  |
| Telephone  | WChar | 50 |  |
| Title  | WChar | 255 |  |
| WWW  | WChar | 255 |  |
| Zip  | WChar | 50 |  |

### AD_REPORT_DATA

| 字段 | 类型 | 长度 | 说明 |
|------|------|------|------|
| CustomerID  | Integer |  | 客户ID |
| DetailID  | Integer |  | →AD_ORDER_DETAILS.PK |
| FoilColour  | WChar | 255 |  |
| LabelPrinted  | Integer |  |  |
| NestANCName  | WChar | 255 |  |
| NestARDName  | WChar | 255 |  |
| NestPartPositionBottom  | Double |  |  |
| NestPartPositionLeft  | Double |  |  |
| NestPartPositionRight  | Double |  |  |
| NestPartPositionTop  | Double |  |  |
| OrderID  | Integer |  | 订单ID |
| PartANCName  | WChar | 255 | NC文件名 |
| PartARDName  | WChar | 255 | 门ARD名 |
| PartItemNumber  | Integer |  | 件号 |
| PartLength  | WChar | 50 |  |
| PartQuantity  | WChar | 50 | 数量 |
| PartQuantityOnSheet  | WChar | 50 | 板内数量 |
| PartStyle  | WChar | 255 | 样式 |
| PartType  | WChar | 255 | 门型 |
| PartWidth  | WChar | 50 |  |
| PathToANC  | WChar | 255 | NC文件路径 |
| PathToARD  | WChar | 255 | 门ARD路径 |
| PathToEMF  | WChar | 255 | 门EMF路径 |
| PathToNestANC  | WChar | 255 |  |
| PathToNestARD  | WChar | 255 |  |
| PathToNestEMF  | WChar | 255 | 整板EMF |
| PathToPressNestARD  | WChar | 255 |  |
| PK 🔑 | Integer |  | 主键 |
| PressDoorCounter  | Integer |  | 门件序号 |
| PressDoorImage  | WChar | 255 | 门标签EMF路径（PressDoorImage） |
| PressItemNumber  | Integer |  |  |
| PressName  | WChar | 255 |  |
| PressPathToEMF  | WChar | 255 |  |
| PressQuantityOnSheet  | Integer |  |  |
| PressQuantityThisSheet  | Integer |  | 压机板内数量 |
| PressSheetIdentifier  | WChar | 50 | 板条码 |
| PressSheetName  | WChar | 50 |  |
| PressSheetNumber  | Integer |  |  |
| ProductionComment  | WChar | 255 | 备注 |
| SheetCount  | WChar | 50 | 嵌套板数 |
| SheetLength  | WChar | 50 |  |
| SheetMaterial  | WChar | 255 | 板材料 |
| SheetName  | WChar | 50 | 板名 |
| SheetNumber  | Integer |  | 板序号 |
| SheetPartCount  | WChar | 50 | 板零件数 |
| SheetScrap  | WChar | 50 | 废料率 |
| SheetThickness  | WChar | 50 |  |
| SheetWidth  | WChar | 50 |  |
| �ͻ�����  | WChar | 50 |  |
| ��������  | WChar | 50 |  |
| ��������  | WChar | 50 |  |

## 三、其他表字段

### AD_DOOR_TREE — 门型树

`ID, NodeDataKey, NodeImage, NodeIndex, NodeKey, NodeParentIndex, NodeTag, NodeTagVariant, NodeText`

### AD_DOOR_XKSET — 门型附加设置

`ID, Project, SubProject, Value1, Value2, Value3, Value4, Value5`

### AD_HANDLE_DRILLING — 手柄钻孔配置（mbln_DrawHandleHoles 读取）

`DatumLocation, DatumLocationAdditionalOffsetX, DatumLocationAdditionalOffsetY, HandleID, HandleName, HoleOrientation, HoleSize, HoleSpacing, MachiningStyle, NumberOfHoles`

### AD_MATERIAL_COLOURS — 材料颜色（膜）配置

`ColourID, ColourName, ColourRotationMethod`

### AD_NEST_ZONES — 排版区域（自动/手动）

`ZoneAssignMethod, ZoneHeight, ZoneID, ZoneName, ZoneNumber, ZonePartArea, ZonePartDimension, ZoneStartPointX, ZoneStartPointY, ZoneWidth`

### AD_OPERATORS — 操作员

`Operator, OperatorDescription, OperatorID`

### AD_ORDER_GRID — 订单表格配置

`DueDate, ID, Jobname, OrderDate, OrderID, PO, ProcessedDate, TotalParts`

### AD_PRESSES — 压机配置

`DefaultPress, GapAtSheetEdge, MinGapBetweenPaths, NumberComponentsBySize, PartRotation, PressID, PressLength, PressName, PressWidth, RectPackTo, TrueShapePackTo, UseTrueShape`

### AD_RULES — 路由器规则（m_PopulateRulesCollection 读取）

`OperatorID, ResultRouter, ResultRouterPost, RuleID, RuleName, RuleText, TestVariableName, TestVariableValue`

### AD_SETTINGS — 系统设置

`NCSeqNum`

### AD_TOOL_ORDER — 刀具排序顺序（mbln_ToolOrdering 读取）

`ToolName, ToolNumber, ToolOffsetNumber, ToolOrderID, ToolSeqNum`

### AD_USER_STYLES — 用户样式宏工程（.arb 路径）

`FullFileName, UserStyleID, VBAProjectName`

### AD_VERSION — 版本信息

`DatabaseVersion, Version`

## 四、数据流（谁写谁读）

| 表 | 写入方 | 读取方 |
|----|--------|--------|
| AD_ORDERS | `modAutoImportNest.glng_CreateOrder`（导入时创建） | `g_Make_Master`/`mstr_GetJobName`/标签同步 |
| AD_ORDER_DETAILS | `modAutoImportNest.ImportCSV`（INSERT...SELECT） | `m_PopulateRouterData`/`m_PopulatePressData`→`m_FillDoorData` |
| AD_DOOR_TYPES | `modAutoImportNest.glng_EnsureStyle`（新门型自动创建） | `mbln_EnsureStyle`/样式加载 |
| AD_DOOR_PATHS | CDM UI 配置 | `mbln_FillPathData`（Make 加工） |
| AD_MATERIALS | `modAutoImportNest.glng_EnsureMaterial`（自动创建） | `m_FillMaterialData` |
| AD_CUSTOMERS | `modAutoImportNest.glng_EnsureCustomer` | 订单装配 |
| AD_REPORT_DATA | `Make.m_InsertReportDataRouter/Press`（排版后） | 报表/标签生成器、`g_RegenDoorLabelEMFs` 同步 |
| AD_TOOL_ORDER | CDM UI | `Make.mbln_ToolOrdering`（刀具排序） |
| AD_RULES | CDM UI | `Make.m_PopulateRulesCollection`/`mbln_RulesApply` |
| AD_NEST_ZONES | CDM UI | `Make.m_PopulateNestZoneCollection`/`m_DrawNestZones` |
| AD_HANDLE_DRILLING | CDM UI | `Make.mbln_DrawHandleHoles` |
| AD_USER_STYLES | CDM UI | `g_GetVBAProjects`（930宏加载） |

## 五、操作注意事项

- 连接：AlphaCAM VBA 用 `gdb_CDM`（`gbln_ConnectToDB()`）；Python 侧无 Access 驱动时经临时宏查询（`run_vba_capture`）
- Jet DAO 的 `LIKE` 通配符默认 `*`（非 `%`）；`INSTR()` 在 Access/SQL Server 均可用
- 直接改库风险：CDM 运行时用内存数据覆盖数据库；改库前备份 `CDM.mdb`（1.2GB）；改后重启 AlphaDOOR 生效
- 删除订单关联：`AD_ORDER_DETAILS`、`AD_REPORT_DATA`（按 OrderID），再删 `AD_ORDERS`（见 `glng_CreateOrder` 覆盖分支）
