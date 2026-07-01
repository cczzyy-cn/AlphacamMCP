# RevNst 新版本分析报告

## 概述

RevNst 是 AlphaCAM ReverseNest 插件中的核心模块（`VB_Name = "RevNst"`），负责在嵌套排料布局中生成**反面（reverse-side）**的镜像副本。该模块包含 3 个公开过程。

## 代码结构

| 过程 | 可见性 | 说明 |
|------|--------|------|
| `g_AroundX` | Public | 绕 X 轴（垂直）镜像 — 在图纸右侧放置镜像线 |
| `g_AroundY` | Public | 绕 Y 轴（水平）镜像 — 在图纸底部放置镜像线 |
| `SetAttribs` | Private | 路径变换工具函数：位移 → 可选反射 → 旋转 → 平移 → 镜像 |

## 功能详解

### 属性常量（Attributes）

模块使用 AlphaCAM 的用户属性机制在几何体和刀路上存储元数据：

| 常量 | 属性名 | 用途 |
|------|--------|------|
| `ATT_PATH_FILE` | `LicomUKsab_nest_path_file` | 零件来源文件名 |
| `ATT_FIRST_PATH` | `LicomUKsab_nest_first_path` | 标记是否为零件的第一条刀路 |
| `ATT_REQUIRED` | `LicomUKsab_nest_required` | 是否为必需零件 |
| `ATT_SHEET_IDENT` | `LicomUKsab_sheet_ident` | 板材标识名 |
| `ATT_SHEET_MATERIAL` | `LicomUKsab_sheet_material` | 板材材质 |
| `ATT_SHEET_THICKNESS` | `LicomUKsab_sheet_thickness` | 板材厚度 |
| `ATT_PART_MOVEX/Y` | `LicomUKsab_part_movex/y` | 零件在嵌套中的 X/Y 平移量 |
| `ATT_PART_ROTANGLE` | `LicomUKsab_part_rotangle` | 零件旋转角度 |
| `ATT_PART_MIRRORED` | `LicomUKsab_part_mirrored` | 零件镜像标记 |
| `ATT_IS_BOBBLE` | `LicomUKsab_is_bobble` | 标识气泡标注圈 |
| `ATT_PART_MOVE_BY_X/Y` | `LicomUKja_part_move_by_x/y` | 额外移动量（SDO 新增） |
| `ATT_PART_SHIFT_X/Y` | `LicomUKja_part_shift_x/y` | 额外偏移量（SDO 新增） |
| `ATT_NEST_ITEM_NUM` | `LicomUKsab_nest_item_number` | 嵌套项目编号 |
| `ATT_IS_REV_SIDE` | `AcamUSrg_IsReverseSide` | 标记为反面（Oct 2010, rg） |
| `ATT_REV_TEXT` | `AcamUSrg_TextIsReversed` | 标记反面文字已处理（Oct 2010, rg） |

### 主流程（g_AroundX / g_AroundY）

两个过程结构相同，区别仅在于镜像轴方向：

#### 阶段 1：板材几何镜像
1. 遍历所有嵌套板材，计算全局最小/最大范围
2. 在图纸边缘外 5% 处建立镜像线
3. 逆序遍历所有几何体（确保拷贝顺序正确）
4. 对每个板材矩形和气泡标注圈：
   - `CopyTemporary` → `MirrorL` → `StoreTemporary`
   - 设置 `ATT_IS_REV_SIDE = 1`
   - 板材名加 `" rev"` 后缀（替代原 `"(reverse)"`）
5. **文字处理**（Oct 2010, rg 新增）：
   - 遍历所有文字对象
   - 检查 `ATT_REV_TEXT = 0`（未处理）
   - 用 `ConvertToTemporaryGeometry` 转几何
   - 用 `TestInsidePath` 检测是否在板材内
   - 在反面位置复制文字
   - 标记 `ATT_REV_TEXT = 1`

#### 阶段 2：刀路和零件镜像
1. 按板材遍历嵌套实例
2. 读取 `ATT_PATH_FILE` 获取零件文件名
3. 构建反面文件名（如 `part_rev.amd`）
4. 尝试 `App.OpenTempDrawing` 加载反面文件
5. 遍历刀路，调用 `SetAttribs` 进行变换
6. 可选（`chkGeos`）：同时复制非刀路的几何体

#### 阶段 3：操作排序
- 支持按板材排序（`bSheetOrder`）
- 支持最小换刀（`bMinToolChanges`）
- 最终调用 `Drw.Operations.OrderAll`

### SetAttribs 函数

路径变换管线（严格顺序）：
1. `MoveL(ShiftX, ShiftY)` — 偏移
2. 若 `intReflect = 1`：`MirrorL(0,1,0,0)` 绕 X 轴反射 + 旋转角取反
3. `RotateL(rotate, 0, 0)` — 旋转
4. `MoveL(MoveX, MoveY)` — 平移到嵌套位置
5. `MirrorL(...)` — 绕镜像线做最终镜像
6. 设置属性：`ATT_FIRST_PATH`、`ATT_PATH_FILE`、`ATT_REQUIRED`、`ATT_NEST_ITEM_NUM`、`ATT_IS_REV_SIDE`

## 与 RevNest_source/modMain.bas 的差异

| 特性 | 粘贴的新版本 (RevNst.bas) | 现有 modMain.bas |
|------|---------------------------|-------------------|
| **SaveSetting 持久化** | ✅ `SaveSetting` 在 `g_AroundX` 和 `g_AroundY` 结束时保存镜像线坐标 | ❌ 无 |
| **Workplane 安全检查 (TFS#80910)** | ❌ 不存在 | ✅ 加载反面文件后检查 `tmpdrw.WorkPlanes.count > 0` 并报错 |
| 其余代码逻辑 | 完全相同 | 完全相同 |

## 关键差异分析

### 1. SaveSetting 持久化
- `g_AroundX` 结束时保存：`mirrorx, mirrorx, miny, maxy`
- `g_AroundY` 结束时保存：`minx, maxx, mirrory, mirrory`
- 注册表路径：`"Alp33082572"` \ `"Holes"` \ `"X1"`, `"X2"`, `"Y1"`, `"Y2"`
- **用途**：保存上次运行的镜像线坐标，供其他工具或流程参考
- **风险等级**：低 — 仅写注册表，不影响核心功能

### 2. Workplane 安全检查缺失
- 现有 `modMain.bas` 在加载反面文件后有一段代码：
  ```vba
  ' TFS#80910 - test to ensure the drawing we are inserting does not contain workplanes
  If tmpdrw.WorkPlanes.count > 0 Then
    MsgBox strName & Chr(13) & gv_CTX(30, 5, "contains workplanes and cannot be used"), vbExclamation
    GoTo loopagain
  End If
  ```
- 此检查阻止包含工作平面的零件被插入（会导致坐标错误）
- **新版本缺失此检查，如在包含工作平面的零件上运行可能导致异常**
- **风险等级**：中 — 建议合并此修复

## 建议

1. **合并工作平面检查**：将 `modMain.bas` 中的 TFS#80910 安全补丁合并到新版本中
2. **添加错误处理**：所有代码在 `On Error Resume Next` 下运行，不会抛出错误但可能静默失败
3. **SaveSetting 改进**：考虑使用更有意义的应用程序名而非 `"Alp33082572"`（可能为许可证哈希）

## 文件保存

- 新模块代码：`RevNest_source/RevNst.bas`（31,719 字节，836 行）
- 本分析文档：`RevNest_source/RevNst_analysis.md`
