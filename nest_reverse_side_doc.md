# Nesting Reverse-Side Mirroring 模块文档

> **来源**: LicomUKsab 插件（Licom UK, Nesting 模块）
> **开发者签名**: SDO（核心逻辑 + 2010-10-01 重构）、rg（2010-10-01 文本处理）
> **文件**: 独立 VBA 模块，约 836 行
> **AlphaCAM API**: VBA (Add-In)

---

## 1. 模块概览

该模块为 **AlphaCAM Nesting 排版** 提供**反面镜像**功能。当用户完成正面排版并生成刀具路径后，本模块自动生成反面（reverse side）的 Sheet 几何、ID 气泡和刀具路径。

### 核心流程

```
正面排版（Nesting） 
    │
    ├─ 1. 计算镜像轴（Sheet 边界外偏移 5%）
    ├─ 2. 镜像 Sheet 几何 + ID 气泡（倒序遍历 Drw.Geometries）
    ├─ 3. 对每个 Sheet 的每个 Part：
    │      ├─ 读取 Part 属性（路径、旋转、位移、反射）
    │      ├─ 打开 *_rev.amd 反面图档
    │      ├─ 对每条 ToolPath 调用 SetAttribs() 应用变换
    │      └─ 设置 OpNo（顺序模式 / 最小换刀模式）
    └─ 4. Operations.OrderAll() 重排操作
```

### 公开接口

| 过程 | 说明 |
|---|---|
| `g_AroundX(bSheetOrder, bMinToolChanges)` | 绕**垂直镜像轴**（X 轴方向，远右端）镜像 |
| `g_AroundY(bSheetOrder, bMinToolChanges)` | 绕**水平镜像轴**（Y 轴方向，远下端）镜像 |
| `SetAttribs(orgPath, curPth, ...)` | 私有辅助，对路径/几何应用 Nest 属性变换 |

---

## 2. 属性常量体系

本模块使用 AlphaCAM **Path.Attribute** 系统存储排版数据。属性命名遵循 `公司_国家_开发者缩写_项目_标识符` 规范。

### 2.1 Nest 属性（LicomUKsab — Nesting）

| 常量 | 值 | 类型 | 说明 | 赋值时机 |
|---|---|---|---|---|
| `ATT_PATH_FILE` | `LicomUKsab_nest_path_file` | String | 部件路径文件名 | Nesting 排版时由 Nest 引擎写入 |
| `ATT_FIRST_PATH` | `LicomUKsab_nest_first_path` | Integer (0/1) | 是否为该部件的首条路径（=1 表示主路径） | Nesting 排版时 |
| `ATT_REQUIRED` | `LicomUKsab_nest_required` | Integer (0/1) | 是否为必须路径 | Nesting 排版时 |
| `ATT_SHEET_IDENT` | `LicomUKsab_sheet_ident` | String | Sheet 标识名称（如 "Sheet1"） | Nesting 排版时 |
| `ATT_SHEET_MATERIAL` | `LicomUKsab_sheet_material` | String | Sheet 材料 | Nesting 排版时 |
| `ATT_SHEET_THICKNESS` | `LicomUKsab_sheet_thickness` | Double | Sheet 厚度 | Nesting 排版时 |
| `ATT_PART_MOVEX` | `LicomUKsab_part_movex` | Double | Part 在 Nesting 中的 X 偏移 | Nesting 排版时 |
| `ATT_PART_MOVEY` | `LicomUKsab_part_movey` | Double | Part 在 Nesting 中的 Y 偏移 | Nesting 排版时 |
| `ATT_PART_ROTANGLE` | `LicomUKsab_part_rotangle` | Double | Part 旋转角度（度） | Nesting 排版时 |
| `ATT_PART_MIRRORED` | `LicomUKsab_part_mirrored` | Integer (0/1) | Part 是否已被反射 | Nesting 排版时 |
| `ATT_IS_BOBBLE` | `LicomUKsab_is_bobble` | Integer (0/1) | 几何是否为 ID 气泡（标识圆） | Nesting 排版时 |
| `ATT_NEST_ITEM_NUM` | `LicomUKsab_nest_item_number` | Integer | Nest 项目编号（2010-10-01 新增） | Nesting 排版时 |

### 2.2 反面属性（AcamUSrg — 用户自定义）

| 常量 | 值 | 类型 | 说明 | 赋值时机 |
|---|---|---|---|---|
| `ATT_IS_REV_SIDE` | `AcamUSrg_IsReverseSide` | Integer (0/1) | =1 表示为反面几何/路径 | 本模块镜像时设置 |
| `ATT_REV_TEXT` | `AcamUSrg_TextIsReversed` | Integer (0/1) | =1 表示该 Text 对象已被处理过镜像 | 本模块镜像时设置 |

### 2.3 附加位移属性（LicomUKja — 另一开发者）

| 常量 | 值 | 类型 | 说明 | 赋值时机 |
|---|---|---|---|---|
| `ATT_PART_MOVE_BY_X` | `LicomUKja_part_move_by_x` | Double | Part 基准 X 偏移 | Nesting 排版时 |
| `ATT_PART_MOVE_BY_Y` | `LicomUKja_part_move_by_y` | Double | Part 基准 Y 偏移 | Nesting 排版时 |
| `ATT_PART_SHIFT_X` | `LicomUKja_part_shift_x` | Double | Part 额外 X 位移 | Nesting 排版时 |
| `ATT_PART_SHIFT_Y` | `LicomUKja_part_shift_y` | Double | Part 额外 Y 位移 | Nesting 排版时 |

> **注意**: `LicomUKja_*` 属性来自不同开发者 "ja"，在 `SetAttribs` 中与 `LicomUKsab_*` 一起使用。先应用 `Shift`（局部偏移），再应用 `MoveBy`（基准偏移），两者作用不同。

---

## 3. 过程签名与参数

### `g_AroundX` / `g_AroundY`

```vba
Public Sub g_AroundX(ByVal bSheetOrder As Boolean, ByVal bMinToolChanges As Boolean)
Public Sub g_AroundY(ByVal bSheetOrder As Boolean, ByVal bMinToolChanges As Boolean)
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `bSheetOrder` | Boolean | `True` = 按 Sheet 分组排列操作（每个 Sheet 的操作移到末尾再插入反面） |
| `bMinToolChanges` | Boolean | `True` = 最小化换刀次数（同类刀具合并到同一操作号）；`False` = 顺序递增 OpNo |

### `SetAttribs`

```vba
Private Function SetAttribs( _
    orgPath As Path,      ' 原始正面路径（从 NestInformation.Parts.Paths(1) 获取）
    curPth As Path,       ' 当前待处理的路径（来自 *_rev.amd 的 ToolPath/Geometry）
    isFirstTP As Long,    ' =1 表示当前路径是部件的首条路径（设置 ATT_FIRST_PATH）
    dblMirror As Double,  ' 镜像轴位置（X 轴坐标或 Y 轴坐标）
    dblMin As Double,     ' 镜像范围最小值
    dblMax As Double,     ' 镜像范围最大值
    intReflect As Integer, ' 输出参数（未使用）
    blnAboutX As Boolean, ' True=绕 X 轴镜像；False=绕 Y 轴镜像
    sName As String       ' 反面文件名（设置 ATT_PATH_FILE）
) As Boolean
```

> **注意**: `SetAttribs` 声明为 `Function` 但未赋值返回值（始终返回 `False`），调用方也未检查返回值。语义上应为 `Sub`。

---

## 4. 核心算法 — Sheet 镜像

### 4.1 镜像轴计算

```vba
' g_AroundX: 垂直镜像线（绕 X 轴方向），放在最左侧 Sheet 的左方 5% 处
mirrorx = minx - ((maxx - minx) * 0.05)

' g_AroundY: 水平镜像线（绕 Y 轴方向），放在最下方 Sheet 的下方 5% 处
mirrory = miny - ((maxy - miny) * 0.05)
```

`minx/maxx/miny/maxy` 通过遍历所有 `ni.Sheets` 的 `sh.Geometry` 边界聚合得到。

### 4.2 倒序遍历几何

```vba
For count = Drw.GetGeoCount To 1 Step -1
```

**倒序原因**: 每次 `StoreTemporary` 会在数据库中插入新路径，倒序遍历确保新插入的对象不会影响尚未处理的原始对象的索引位置。

### 4.3 Sheet 几何镜像流程

```
对每个 Sheet/Dimension 几何:
  │
  ├─ 跳过无 ATT_IS_BOBBLE 属性的 Dimension 对象
  │
  ├─ CopyTemporary → MirrorL → StoreTemporary
  │
  ├─ 若是 Sheet:
  │     ├─ 设置 ATT_SHEET_IDENT = 原名 + " rev"
  │     ├─ 复制材料/厚度属性
  │     ├─ 设置 ATT_IS_REV_SIDE = 1
  │     ├─ 遍历 Drw.Text:
  │     │     ├─ 跳过已处理的 (ATT_REV_TEXT = 1)
  │     │     ├─ ConvertToTemporaryGeometry
  │     │     ├─ TestInsidePath(原始Sheet) → 确认属于该 Sheet
  │     │     ├─ 镜像临时几何
  │     │     ├─ MoveL 将文字复位（X方向或Y方向偏移补偿）
  │     │     ├─ 复制 Text 对象，设置 ATT_REV_TEXT = 1
  │     │     └─ 更新 ATT_SHEET_IDENT
  │     └─ 分配 Group 编号
  │
  └─ 若是 ID 气泡（Dimension + IsArc）:
        ├─ LCase(Sheet名) 作为气泡内文字
        ├─ 创建临时文本测量尺寸
        ├─ 计算缩放因子 SF = 0.707 * 气泡高 / 文字最大边
        ├─ 删除临时文本
        ├─ 计算居中位置
        └─ 创建最终文本并设 Group + Dimension=True
```

### 4.4 ID 气泡文本处理详解

该段代码在 Sheet 镜像后对 ID 气泡（标识圆）内部的文字进行处理：

1. 用 `LCase(wrd)` 将 Sheet 名转为小写
2. 创建临时文本来测量文字尺寸（`Drw.CreateText(wrd, 0, 0, high)`）
3. 计算文字外接矩形，取 `Max(宽, 高)` 作为 `big`
4. 缩放因子 `SF = 0.707 * 气泡高度 / big`（确保文字不超过气泡直径的 70.7%）
5. 文字居中位置 = 气泡中心 - (文字尺寸/2 * SF)
6. 最终文字高度 = 气泡高度 * SF
7. 设置 `tp.Group = grp` 和 `tp.Dimension = True`（标记为标注，不会再次被处理）

### 4.5 Sheet 排序模式

当 `bSheetOrder = True` 时，在插入反面路径前先对当前 Sheet 的所有操作重编号：

```vba
' 找到当前 Sheet 的最大和最小 OpNo
For Each tp In sh.Paths
    If maxop < tp.OpNo Then maxop = tp.OpNo
    If minop > tp.OpNo Then minop = tp.OpNo
Next tp

' 将所有操作移到尾部
For I = minop To maxop
    Drw.Operations.Renumber minop, lastop, acamOpADD_TO_OPERATION
Next I
```

效果：同一个 Sheet 的操作在操作列表中连续排列，反面路径紧随其后。

---

## 5. 核心算法 — ToolPath 镜像

### 5.1 反面文件加载机制

```vba
strName = P.Attribute(ATT_PATH_FILE)           ' 如 "C:\parts\door_left.amd"
prefix = Left(strName, Len(strName) - 4)       ' "...door_left"
suffix = Right(strName, 4)                     ' ".amd"
strName = prefix + "_rev" + suffix             ' "...door_left_rev.amd"
Set tmpdrw = App.OpenTempDrawing(strName)      ' 打开反面图档（临时）
```

**命名规则**: 在原文件名扩展名前插入 `_rev` 后缀。如 `door_left.amd` → `door_left_rev.amd`。如果文件不存在，`OpenTempDrawing` 返回 `Nothing`，该 Part 被跳过。

### 5.2 ToolPath 遍历与变换

```
对每个 Part (inst):
  │
  ├─ 从 inst.Paths(1) 获取正面首条路径
  │
  ├─ 构造 _rev 文件名并 OpenTempDrawing
  │
  ├─ 对反面图档中的每条 ToolPath:
  │     ├─ SetAttribs() 应用属性变换
  │     ├─ CopyTemporary → 设置 OpNo → StoreTemporary
  │     └─ 更新 OpNo 计数器
  │
  └─ [可选] 对反面图档中的普通几何 (chkGeos):
        └─ SetAttribs() → CopyTemporary → StoreTemporary
```

### 5.3 OpNo 分配策略

#### 顺序模式（`bMinToolChanges = False`）

```
pcopy.OpNo = lastop
lastop = lastop + 1
```

每条路径分配唯一的递增操作号。简单直接，但可能导致多次换刀。

#### 最小换刀模式（`bMinToolChanges = True`）

```vba
Set currtool = pcopy.GetTool
For I = sheetop To lastop - 1
    If currtool.Number = Drw.Operations(I).Tool.Number Then Exit For
Next I
pcopy.OpNo = I
If I = lastop Then lastop = lastop + 1
```

在已添加的反面路径中寻找同号刀具的操作，将当前路径归入同一操作号。若未找到则新建操作号。

### 5.4 `SetAttribs` 变换顺序（关键！）

```
初始状态（从 *_rev.amd 加载的路径）
    │
    ├─ 1. MoveL(ShiftX, ShiftY)          ← 局部偏移（ATT_PART_SHIFT_*）
    │
    ├─ 2. [若 ATT_PART_MIRRORED=1]:
    │      └─ MirrorL(0,1, 0,0)          ← 水平反射（绕 X 轴翻转）
    │      └─ rotate = -rotate           ← 反射后旋转方向取反
    │
    ├─ 3. RotateL(rotate, 0, 0)          ← 旋转（ATT_PART_ROTANGLE，绕原点）
    │
    ├─ 4. MoveL(MoveX, MoveY)            ← 基准偏移（ATT_PART_MOVE_BY_*）
    │
    └─ 5. MirrorL(镜像轴, 范围)           ← 绕全局镜像轴反射
         ├─ blnAboutX=True:  MirrorL(mirrorX, minY, mirrorX, maxY)
         └─ blnAboutX=False: MirrorL(minX, mirrorY, maxX, mirrorY)
```

**变换顺序的重要性**: 
- Shift（局部偏移）在反射/旋转前应用，确保部件的**内部几何关系**正确
- 反射在旋转前（若 Part 已被反射），旋转方向取反
- MoveBy（全局定位）在反射/旋转后应用
- 最后绕全局镜像轴反射，将路径从正面位置映射到反面位置

### 5.5 属性传递

反面路径的属性设置：

| 属性 | 值来源 |
|---|---|
| `ATT_FIRST_PATH` | `isFirstTP`（首条为 1，其余为 0） |
| `ATT_PATH_FILE` | 新构造的 `_rev` 文件名 |
| `ATT_REQUIRED` | 从 `orgPath` 复制 |
| `ATT_NEST_ITEM_NUM` | 从 `orgPath` 复制 |
| `ATT_IS_REV_SIDE` | = 1（固定标识为反面） |

---

## 6. 注册表存储

```vba
' g_AroundX:
SaveSetting "Alp33082572", "Holes", "X1", mirrorx
SaveSetting "Alp33082572", "Holes", "X2", mirrorx
SaveSetting "Alp33082572", "Holes", "Y1", miny
SaveSetting "Alp33082572", "Holes", "Y2", maxy

' g_AroundY:
SaveSetting "Alp33082572", "Holes", "X1", minx
SaveSetting "Alp33082572", "Holes", "X2", maxx
SaveSetting "Alp33082572", "Holes", "Y1", mirrory
SaveSetting "Alp33082572", "Holes", "Y2", mirrory
```

将镜像轴范围写入注册表（HKCU\Software\VB and VBA Program Settings\Alp33082572\Holes），可能用于后续的钻孔或后处理参考。

---

## 7. 外部依赖

| 依赖 | 类型 | 说明 |
|---|---|---|
| `g_LockAcam` / `g_UnlockAcam` | 全局过程 | 操作锁定/解锁 AlphaCAM 界面，防止操作冲突 |
| `gv_CTX(ctxID, msgID, default)` | 全局函数 | 多语言字符串获取（上下文 30 = "Reverse" 相关） |
| `frmMain.chkGeos` | UserForm 控件 | 复选框，控制是否同时镜像普通几何（非 ToolPath） |
| `App.OpenTempDrawing(fileName)` | AlphaCAM API | 打开图档为临时绘图（不添加到 MRU 列表） |
| `App.ActiveDrawing.GetNestInformation` | AlphaCAM API | 获取 NestInformation 对象，遍历 Sheet/Part |
| `NestInformation` / `NestSheet` / `NestPartInstance` | AlphaCAM API | Nesting 排版信息对象层次结构 |
| `Path.CopyTemporary` / `Path.StoreTemporary` | AlphaCAM API | 复制路径到临时缓冲区，再存入数据库 |
| `Path.MirrorL` / `Path.RotateL` / `Path.MoveL` | AlphaCAM API | 2D Local 坐标系的几何变换 |
| `Operations.Renumber` | AlphaCAM API | 重编号操作（`acamOpADD_TO_OPERATION` 标志追加到目标操作） |
| `Operations.OrderAll` | AlphaCAM API | 按 OpNo 对所有路径排序 |

---

## 8. 潜在问题与架构笔记

### 8.1 代码质量问题

| 问题 | 位置 | 说明 |
|---|---|---|
| `On Error Resume Next` | 所有过程顶部 | 全局吞没所有运行时错误，调试困难。若出错，镜像可能部分完成但无声 |
| 密集 GoTo 跳转 | 每过程 3-4 个标签 | `byebye` / `loopnext` / `loopagain` / `loopalso` 使控制流不清晰 |
| `g_AroundX` / `g_AroundY` 高度重复 | 两个过程约 350 行 | 只有镜像轴计算和 `MirrorL` 参数不同，其余 ~90% 相同 |
| `SetAttribs` 不返回值 | Function 声明 | 已声明为 `Function` 但无 `SetAttribs = True/False`，调用方也不检查 |
| 硬编码注册表键 | `SaveSetting "Alp33082572"` | 无常量定义，`Alp33082572` 含义不明 |
| 临时绘图未显式关闭 | `OpenTempDrawing` 循环 | `tmpdrw` 变量被覆盖时未调用 `Close`，可能导致内存泄漏 |
| 无错误恢复 | 所有过程 | 出错后掉入 `byebye` 清理段，但无法区分是正常完成还是出错退出 |

### 8.2 架构笔记

- **`LicomUKsab` vs `LicomUKja` 属性共存**: 两套位移属性（`MoveBy/Shift` 来自 ja，`MoveX/MoveY` 来自 sab）同时在 `SetAttribs` 中使用，先 Shift 后 MoveBy。这可能反映了不同版本 Nest 引擎的兼容。
- **`isFirstTP` 标志的传递**: `SetAttribs` 在每个 Part 的首条路径设置 `ATT_FIRST_PATH=1`，后续为 0。但 `isFirstTP` 参数在循环外声明为 `Long`，循环中首次后设为 0 — 这意味着如果 `OpenTempDrawing` 失败（文件不存在），`isFirstTP` 不会递增，同一个 Part 的下一次尝试仍视为首条。这是合理的行为，因为失败不会创建任何路径。
- **2010-10-01 重构**: 代码注释中的 `+SDO` 和 `-rg` 标记表明该日期有较大重构：
  - SDO 将路径变换逻辑抽入 `SetAttribs` 函数（此前内联在 `g_AroundX` 和 `g_AroundY` 中）
  - SDO 添加几何镜像支持（`chkGeos`）
  - rg 将反面后缀从 `(reverse)` 改为 ` rev`，并修改了 ID 气泡文字处理逻辑
  - rg 添加了 `ATT_IS_REV_SIDE` 和 `ATT_REV_TEXT` 属性体系
- **NestInformation.Paths(1) 约定**: 代码假定 `inst.Paths(1)` 是带有完整属性（特别是 `ATT_PATH_FILE`、`ATT_PART_MOVEX/Y` 等）的首条路径。这一约定依赖于 Nest 引擎在排版时始终将首条路径设为主要属性载体。
- **`chkGeos` 的访问时机**: `frmMain.chkGeos.Value` 在循环中每次都被读取，因此用户可以在镜像过程中更改复选框状态并影响后续 Sheet 的处理。

---

## 9. 开发者签名注释

```
'01 OCT 10  +SDO               ← SDO: 添加 ATT_NEST_ITEM_NUM / ATT_IS_REV_SIDE
'01 OCT 10 +-SDO               ← SDO: 将路径操作抽入 SetAttribs，添加几何镜像（chkGeos）
' 01 oct 10 - rg               ← rg: 修改文字处理、后缀名、属性体系
```

**推断**:
- **SDO** — 主要开发者（Stephen A.B. Jones？），完成初始版本和 2010-10-01 重大重构
- **rg** — 后期贡献者，2010-10-01 修改文字处理和反面标识体系
- **ja** — `LicomUKja_*` 属性作者，可能为另一 Nest 组件开发者

---

## 附录 A: AlphaCAM API 参考

| API 调用 | 说明 |
|---|---|
| `Path.MirrorL(x1, y1, x2, y2)` | 绕 Local 坐标系中通过两点的直线镜像 |
| `Path.RotateL(angle, cx, cy)` | 在 Local 坐标系中绕指定中心旋转（角度单位为度） |
| `Path.MoveL(dx, dy)` | 在 Local 坐标系中平移 |
| `Path.MoveG(dx, dy, dz)` | 在 Global 坐标系中 3D 平移 |
| `Path.CopyTemporary` | 复制路径但不加入数据库 |
| `Path.StoreTemporary` | 将临时路径写入数据库 |
| `Path.Attribute(name)` | 读写自定义属性（Variant 类型） |
| `Path.TestInsidePath(path)` | 测试路径是否在另一封闭路径内部（返回 acamResultTRUE/FALSE） |
| `Text.ConvertToTemporaryGeometry` | 将 Text 对象转换为临时几何路径 |
| `Drw.GetNestInformation` | 返回 NestInformation 对象 |
| `Operations.Renumber(old, new, flag)` | 重编号操作，`acamOpADD_TO_OPERATION` 追加到目标操作 |

---

## 10. 模块全景 — 全部 8 个模块分析

> 以下分析基于从 `ReverseNest.amb`（OLE2 复合文档）使用 olevba 提取的反编译源码。
> 源码文件已提取至 `RevNest_source/` 目录。

### 10.1 模块清单

| # | 模块 | 大小 | 类型 | 职责 |
|---|---|---|---|---|
| 1 | `modGlobal.bas` | 3 KB | 标准模块 | 全局常量/枚举/UDT 声明 |
| 2 | `Events.bas` | 2 KB | 类模块 | 插件入口，菜单注册/更新 |
| 3 | `frmMain.frm` | 3 KB | UserForm | 选项对话框 |
| 4 | `modMain.bas` | 32 KB | 标准模块 | 核心反面镜像逻辑（g_AroundX/Y） |
| 5 | `modAcam.bas` | 140 KB | 标准模块 | AlphaCAM API 封装库（70+ 函数） |
| 6 | `modRegistry.bas` | 21 KB | 标准模块 | 注册表读写/枚举/导入导出 |
| 7 | `modFilesFolders.bas` | 31 KB | 标准模块 | 文件/路径/特殊文件夹工具 |
| 8 | `modGeneral.bas` | 33 KB | 标准模块 | UI/字符串/数字通用工具 |

### 10.2 模块间依赖图

```
modGlobal (常量/枚举/类型)
    ↑
    ├──→ modRegistry (注册表 API)
    ├──→ modFilesFolders (文件/路径工具)
    ├──→ modGeneral (UI/字符串工具)
    ├──→ modAcam (AlphaCAM 核心库) ← 使用 modRegistry, modFilesFolders, modGeneral
    │
    ├──→ Events (插件入口) ← 使用 modAcam, modFilesFolders, frmMain
    │
    ├──→ frmMain (对话框) ← 使用 modMain, modAcam, modRegistry, modGeneral
    │
    └──→ modMain (执行引擎) ← 使用 modAcam, frmMain
```

### 10.3 运行时调用链

```
AlphaCAM 启动 → InitAlphacamAddIn (Events.bas)
                    ↓
用户点击菜单 → g_RevSide (Events.bas)
                    ↓
            → frmMain.Show (对话框弹出)
                    ↓
用户点击 OK → cmdOK_Click (frmMain)
                    ↓
            → g_AroundX / g_AroundY (modMain.bas)
                    ↓
            → g_LockAcam · gv_CTX · NestInformation · etc. (modAcam)
            → g_UnlockAcam
```

### 10.4 各模块详细分析

#### 10.4.1 `modGlobal.bas` — 全局声明

仅声明，无可执行代码。

| 声明 | 说明 |
|---|---|
| `DEF_APP_TITLE` = `"Reverse Side Nesting"` | 应用标题 |
| `DEF_VERSION` = `"1.2"` | 版本号 |
| `DEF_MACRO_NAME` = `"ReverseNest"` | 宏项目名 |
| `DEF_CTX` = `"ReverseNest.txt"` | 语言文件 |
| `AlphaIntersectPoint` 枚举 | 线线相交结果 |
| `POINT_XYZ`, `WP_XYZ`, `LINE_XYZ`, `ARC_DETAILS` UDT | 几何数据类型 |
| `LicomUKsab_*`, `LicomUKja_*`, `LicomUKjba_*` | Nest 属性前缀常量 |

#### 10.4.2 `Events.bas` — 插件入口

| 函数 | 说明 |
|---|---|
| `InitAlphacamAddIn(AcamVersion)` | 注册菜单项，仅当 `gi_NestLevel = 1`（Nesting 层级）时生效 |
| `g_RevSide()` | 菜单点击处理：加载并显示 frmMain |
| `OnUpdateg_RevSide()` | 菜单状态更新：当 `gb_HasNesting()` 为真时启用 |

菜单注册代码：
```vba
' 原始版本 (被注释):
' fr.AddMenuItem2 App.Frame.ReadTextFile("ReverseNest.txt", 1, 1), "RevSide", acamMenuSPECIALFUN, ""
'
' 2010-06-10 rg 修改版:
sNestTitle = gs_ReadAcamNestCTX(1140, 1, "&Nesting")
.AddMenuItem3 gv_CTX(1, 1, "Reverse-Side Nesting"), "g_RevSide", acamMenuUTILS_NEST, "", sNestTitle
```

#### 10.4.3 `frmMain.frm` — 选项对话框

**控件布局：**
- `fraSheetOrder`：Sheet 排序 — `optSheetOrderBySide` / `optSheetOrderBySheet`
- `fraFlip`：镜像轴 — `optFlipX` / `optFlipY`
- `chkMinChanges`：最小化换刀
- `chkGeos`：包含几何（2010-10-01 SDO 添加）

**事件：**
- `UserForm_Initialize()`：从 `ReverseNest.txt` 加载本地化标题，从注册表恢复上次设置
- `cmdOK_Click()`：读取选项 → 调用 `g_AroundX` 或 `g_AroundY` → 保存设置到注册表
- `cmdCancel_Click()`：隐藏并卸载窗体

**注册表保存键：**
```
HKCU\Software\Planit\Alphacam\ReverseNest\
    OrderBySheet    (0/1)
    FlipAroundX     (0/1)
    MinToolChanges  (0/1)
    IncludeGeos     (0/1)
```

**旧版兼容：** 使用 `GetSetting("LicomSystems", "_ReverseNest", ...)` 作为默认回退。

#### 10.4.4 `modMain.bas` — 核心镜像逻辑

包含两个主要过程 `g_AroundX` 和 `g_AroundY`（约 350 行各），以及辅助函数 `SetAttribs`。

**常量体系（14 个 Private Const）：** 参见第 2 节常量表。

**算法流程：** 参见第 4-5 节详细分析。

**与用户粘贴代码的差异：** 从 .amb 反编译的源码包含额外的工作平面检查（TFS#80910）：
```vba
' TFS#80910 - test to ensure the drawing we are inserting does not contain workplanes
If tmpdrw.WorkPlanes.count > 0 Then
    MsgBox strName & Chr(13) & gv_CTX(30, 5, "contains workplanes and cannot be used"), vbExclamation
    GoTo loopagain
End If
```

#### 10.4.5 `modAcam.bas` — AlphaCAM 核心库（140 KB，最大模块）

通用 AlphaCAM API 封装，包含约 **70+ 个 Public 函数**。大部分与反面镜像无直接关联，属于复用工具库。

**枚举（11 个）：**
- `AlphaModuleType` — 模块类型（Wire/Profiling/Mill/Router/Stone/Lathe）
- `AlphaVariableType` — 变量类型
- `AlphaFileType` — 文件类型
- `AlphaExtensionType` — 文件扩展名类型
- `AlphaWorkFace` — 工作面（Top/Front/Right/Back/Left/Bottom）
- `AlphaNestExtension` — Nest 扩展功能
- 等

**关键 Public 函数分类：**

| 类别 | 函数 |
|---|---|
| **加解锁** | `g_LockAcam()` / `g_UnlockAcam(bZoomAll)` — 启用/禁用屏幕重绘 |
| **本地化字符串** | `gv_CTX(dollar, index, default)` — 读取 `ReverseNest.txt` 语言文件 |
| **Nest 检测** | `gb_HasNesting()` / `gi_NestLevel()` / `gb_HasNestExtension()` |
| **几何变换** | `g_LtoG()` / `g_GtoL()` — 局部/全局坐标转换；`gp_Offset()` — 偏移 |
| **几何查询** | `gb_IsCircle()` / `gb_IsRectangle()` / `gb_IsLine()` |
| **属性管理** | `g_CopyObjectAtts()` / `g_CopyPathAtts()` / `g_ClearAttributes()` / `g_WipeNestAttributesFromPath()` / `g_PrintAttributsToDebug()` |
| **层管理** | `glyr_GetLayer()` / `glyr_GetActiveLayer()` / `g_HideAllUserLayers()` |
| **文件路径** | `gs_LICOMDAT()` / `gs_LICOMDIR()` / `gs_AppDir()` / `gs_TemplateDir()` |
| **绘图操作** | `gb_SaveDrawing()` / `gb_StartFileNew()` / `gb_OutputNC()` |
| **后处理** | `gb_PostVariableExists()` / `gl_GetAvailablePostAttNumber()` |
| **视图** | `g_Redraw()` / `g_HideAllToolpaths()` / `g_ShowAllGeos()` |
| **样式/刀具** | `gb_StyleExists()` / `gb_PickTool()` / `gb_OpenTool()` |
| **工作平面** | `g_GetWAAandWACandWTC()` / `gi_GetWVF()` / `gb_IsSameWP()` |
| **碰撞检测** | `gb_PathsAreOnSamePlane()` |

#### 10.4.6 `modRegistry.bas` — 注册表工具（21 KB）

完整的注册表操作层：

| 函数 | 说明 |
|---|---|
| `gs_ReadRegKey(keyPath, subKey, root, default)` | 读取注册表值 |
| `gb_WriteRegKey(type, keyPath, subKey, value, root)` | 写入（REG_SZ / DWORD / BINARY） |
| `gs_EnumerateRegKeys(root, keyPath)` | 枚举子键 |
| `gs_EnumerateRegKeyValues(root, keyPath)` | 枚举键值 |
| `gb_DeleteRegKey(root, keyPath, subKey)` | 删除键 |
| `gb_DeleteRegKeyValue(root, keyPath, subKey)` | 删除值 |
| `gb_ExportRegKey(root, keyPath, fileName)` | 导出键到 .reg 文件 |
| `gb_ImportRegKey(root, keyPath, fileName)` | 从 .reg 文件导入 |

**默认注册表路径：** `HKCU\Software\Planit\Alphacam\ReverseNest\`

**安全提升：** `mb_EnablePrivilege()` — 为导入导出操作提升 SE_BACKUP_NAME / SE_RESTORE_NAME 权限。

#### 10.4.7 `modFilesFolders.bas` — 文件/路径工具（31 KB）

| 函数 | 说明 |
|---|---|
| `gs_ThisDir()` | 返回 ReverseNest 插件所在目录（通过遍历 VBE 项目定位） |
| `gs_ThisFile()` | 返回 .amb 文件完整路径 |
| `gs_ReadFileContents(file)` | 读取文本文件全部内容 |
| `gs_ParseFileName()` | 从路径中提取文件名 |
| `gs_ParseDirName()` | 提取父目录 |
| `gs_ParseFileExtension()` | 提取扩展名 |
| `gs_ReplaceFileExtension()` | 替换扩展名 |
| `gs_GetLocalAppDataDir()` | 返回 `AppData\Local\Planit\Alphacam\ReverseNest\` |
| `gs_GetCommonAppDataDir()` | 返回 `AppData\Roaming\Planit\Alphacam\ReverseNest\` |
| `gs_GetSpecialFolder()` | 通过 SHGetFolderPath 获取特殊文件夹 |
| `gs_GetDir()` | SHBrowseForFolder 文件夹选择对话框 |
| `gs_UniqueFileName()` | 生成不重复的文件名（追加 (1), (2) ...） |
| `gb_ProjectExists()` | 检查 VBA 项目是否已加载 |
| `gs_MacroDir()` | 返回指定 VBA 项目的目录 |
| `gl_FileOpenSaveDialogCallback()` | 将文件对话框居中 |

#### 10.4.8 `modGeneral.bas` — 通用工具（33 KB）

| 函数 | 说明 |
|---|---|
| `g_SetAccelerators(Frm)` | 遍历窗体控件，设置快捷键 (&) 和区域字体 |
| `g_SetProperFont(Obj, size)` | 根据系统语言设置字体字符集（中文/日文/韩文/希伯来文） |
| `gs_GetRegionalSetting()` | 读取系统区域设置（小数分隔符、日期格式等） |
| `g_Eval()` / `g_EvalInt()` | 文本框数学表达式求值（调用 App.Frame.Evaluate） |
| `gb_CheckAllText()` | 验证所有文本框非空 |
| `gcol_DelimitedStringToCollection()` | 分隔字符串 → Collection |
| `gv_Split()` | 自定义 Split 函数 |
| `gs_RemoveIllegalChars()` | 移除非法文件名字符 `\ / : * ? < > | "` |
| `gs_GUID()` | 生成 GUID 字符串 |
| `g_Repaint()` | 通过 RedrawWindow API 强制重绘窗体 |
| `gl_FrmHwnd()` | 查找 UserForm 的窗口句柄 |
| `g_Help()` | 启动 HTML Help |
| `g_DebugNote()` | 输出调试信息（OutputDebugString） |

---

## 附录 B: 文件清单与提取记录

### ReverseNest 插件文件

| 文件 | 大小 | 说明 |
|---|---|---|
| `ReverseNest.arb` | 415 KB | 编译后的 Add-In（AlphaCAM 运行时加载） |
| `ReverseNest.amb` | 415 KB | 源码/编译混合（OLE2 复合格式，含 VBA 源码流） |
| `ReverseNest.asb` | 415 KB | 二进制源码备份 |
| `ReverseNest.doc` | 81 KB | Word 格式文档（使用说明/命名规则） |
| `ReverseNest.txt` | 1.2 KB | 语言/描述文件（`$1`=`Reverse-Side Nesting`） |
| `ReverseNest.eng` | 1.2 KB | 英文语言文件（同 txt） |

### 提取的源码文件

| 路径 | 说明 |
|---|---|
| `RevNest_source/Events.bas` | 插件入口，菜单注册 |
| `RevNest_source/frmMain.frm` | 选项对话框 |
| `RevNest_source/modMain.bas` | 核心镜像逻辑（32 KB） |
| `RevNest_source/modAcam.bas` | AlphaCAM API 封装（140 KB，最大模块） |
| `RevNest_source/modGeneral.bas` | UI/字符串通用工具（33 KB） |
| `RevNest_source/modFilesFolders.bas` | 文件/路径工具（31 KB） |
| `RevNest_source/modRegistry.bas` | 注册表工具（21 KB） |
| `RevNest_source/modGlobal.bas` | 全局常量/类型（3 KB） |

---

*文档生成日期: 2026-06-19*
*基于对 836 行 VBA 源代码（g_AroundX / g_AroundY / SetAttribs）以及从 ReverseNest.amb 反编译的全部 8 个模块的逆向分析*
