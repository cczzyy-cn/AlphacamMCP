# RevNest (Reverse-Side Nesting) v1.2 API 参考

> 来源: AlphaCAM 2016 R1 插件 `ReverseNest.arb`
> 作者: Licom/Planit (SDO, rg, ja)
> 路径: `...\000\StartUp\Utils\ReverseNest\`

---

## 一、入口与菜单 (Events.bas)

| 函数 | 说明 |
|---|---|
| `InitAlphacamAddIn(AcamVersion)` | 插件入口。注册菜单到 Nesting 选项卡 |
| `g_RevSide()` | 菜单点击处理：弹出 frmMain 对话框 |
| `OnUpdateg_RevSide()` | 菜单状态更新：有排版时启用 |

## 二、对话框 (frmMain.frm)

| 事件 | 说明 |
|---|---|
| `UserForm_Initialize()` | 加载本地化文字，从注册表恢复设置 |
| `cmdOK_Click()` | 读取选项 → 调用 g_AroundX/g_AroundY → 保存设置 |
| `cmdCancel_Click()` | 关闭对话框 |

**控件:** optFlipX/Y(镜像轴), optSheetOrder(排序), chkMinChanges(最小换刀), chkGeos(包含几何)

## 三、核心算法 (modMain.bas)

| 过程 | 说明 |
|---|---|
| `g_AroundX(bSheetOrder, bMinToolChanges)` | 绕垂直轴镜像反面排版 |
| `g_AroundY(bSheetOrder, bMinToolChanges)` | 绕水平轴镜像反面排版 |
| `SetAttribs(orgPath, curPth, ...)` | 私有辅助：对路径应用变换和属性复制 |

## 四、AlphaCAM 工具库 (modAcam.bas) — 70+ 函数

### 加解锁与重绘
`g_LockAcam()` — 禁用屏幕重绘
`g_UnlockAcam(bZoomAll)` — 恢复屏幕重绘
`g_Redraw(bZoomAll, bRefresh)` — 重绘图纸

### Nesting 排版
`gi_NestLevel()` — 返回 Nesting 层级(0/1)
`gb_HasNesting(o)` — 检查是否有排版
`gb_HasSheetGeo()` — 检查 Sheet 几何
`gb_HasNestExtension(name, NE, idx)` — 检查排版扩展功能
`g_SetNestExtension(ext, active, ...)` — 启用/禁用排版扩展
`g_GetMinMaxOpNumbersForSheet(min, max, paths)` — Sheet 操作号范围
`g_TransposeNestedSheetXY(quad, ...)` — 转置 Sheet XY
`g_WipeNestList(N, name)` — 清除排版列表

### 几何查询
`gb_IsCircle(P)` — 是否为圆
`gb_IsRectangle(P)` — 是否为矩形
`gb_IsLine(P)` — 是否为直线
`gb_HasWorkVolume(p)` — 是否有工作体积
`gb_HasMaterial(p)` — 是否有材料块
`gb_HasSolids(sf)` — 是否有实体
`gb_HasAnything()` — 是否有任何路径
`gb_HasAnythingOnWorkPlane(WP)` — 工作平面是否有路径
`gb_HasVisibleToolpaths()` — 是否有可见刀具路径
`gb_PathsAreOnSamePlane(P1, P2)` — 路径共面检测
`gb_IsMultidrillOp(Op)` — 是否为多孔钻
`gp_GetWV()` — 获取工作体积路径
`gb_IsAcamColor(c)` — 验证 AlphaCAM 颜色

### 坐标变换
`g_LtoG(WP, X, Y, Z)` — 局部→全局坐标
`g_GtoL(WP, X, Y, Z)` — 全局→局部坐标
`g_GetExtentsG(PS, min, max)` — 路径集合全局边界
`g_GetWAAandWACandWTC(WP, ...)` — 工作平面旋转角度
`gi_GetWVF(WP)` — 工作面枚举值
`gb_IsSameWP(WP1, WP2)` — 比较工作平面

### 路径操作
`gp_Offset(P, dist, side, delOrig)` — 轮廓偏置
`gps_Array(paths, ...)` — 二维阵列
`gps_Repeat(paths, ...)` — 行/列复制
`g_ExtendByDistance(P, dist, isPct, end)` — 延长/缩短
`g_AppendPathsToCollection(src, dst)` — 合并路径集合
`g_RemovePathFromColleciton(coll, path)` — 移除路径
`gps_GetAllPaths()` — 返回所有路径
`gps_GetPathsInGroup(group, ps)` — 组内路径
`gps_GetPathsNotInGroup(group, ps)` — 组外路径
`gps_GetToolPathsInOperation(Op)` — 操作内刀具路径

### 文本转路径
`gps_ConvertTextToGeometry(TS, bLeaveOrig)` — Text→几何路径

### 属性管理
`g_CopyObjectAtts(from, to, overwrite)` — 复制对象属性
`g_CopyPathAtts(from, to, overwrite, inclElems)` — 复制路径属性
`g_CopyElementAtts(from, to, overwrite)` — 复制元素属性
`g_ClearAttributes(obj)` — 清除所有用户属性
`g_PrintAttributsToDebug(obj, label)` — 输出属性到调试窗口
`gb_AttrExists(attr, obj, val)` — 检查属性是否存在
`gb_AttrExistsInPost(attr, num)` — 检查后处理器属性
`gl_GetAvailablePostAttNumber()` — 可用后处理器属性编号
`g_WipeNestAttributesFromPath(P)` — 清除 Nest 属性

### 图层管理
`glyr_GetActiveLayer()` — 当前图层
`glyr_GetLayer(name, bCreate, ...)` — 获取/创建图层
`gb_IsAlphacamLayer(Lyr)` — 是否为系统层
`gb_IsMachineOrClampLayer(Lyr)` — 是否为机床/夹具层
`g_ShowAllUserLayers()` / `g_HideAllUserLayers()` — 显示/隐藏用户层
`g_HideAllToolpaths()` / `g_ShowAllToolpaths()` — 隐藏/显示刀具路径
`g_HideGeos()` / `g_UnhideGeos()` — 隐藏/取消隐藏几何
`g_ShowAllGeos()` / `g_ShowAllOpGeos()` — 显示几何
`gb_OffsetToLayer()` / `gb_CopyToLayer()` / `gb_MoveToLayer()` — 层操作

### 几何可见性
`g_DisableGeos()` / `g_EnableGeos()` — 禁用/启用几何选择
`g_DisableAllGeosExcept()` / `g_EnableAllGeosExcept()` — 批量控制

### 文件与路径
`gs_LICOMDAT()` / `gs_LICOMDIR()` — 返回系统目录
`gs_TemplateDir()` / `gs_ToolsDir()` / `gs_PostDir()` / `gs_StylesDir()`
`gs_RawModuleName()` — 插件原始文件名
`gl_AcamHwnd()` — AlphaCAM 主窗口句柄

### 文件操作
`gb_StartFileNew(bForce)` — 新建图纸
`gb_SaveDrawing(bForce)` — 保存图纸
`gb_IsDrawingSaved(name)` — 检查是否保存
`gb_OutputNC(file, outputTo, ...)` — 输出 NC
`gb_PostVariableExists(name)` — 后处理器变量是否存在
`gb_HasSTART(file)` — NC 文件是否含 START

### 刀具与样式
`gb_PickTool(tool, MT)` — 交互选刀
`gb_OpenTool(tool, MT)` — 打开刀具文件
`gb_StyleExists(style, msRet)` — 检查加工样式
`gs_GetOperationType(MD, icon)` — 操作类型名称
`go_GetSawMillDataFromPath(P)` — 锯切参数

### 层级与模块
`gi_ModuleType()` — 模块类型(Mill/Router)
`gi_ModuleLevel()` — 程序级别

### Z 层级
`gb_SetGeometryZLevel(P, top, bottom)` — 设置 Z 层级
`gb_SetGeometryZLevelsMultiple(PS, ...)` — 批量设置
`gb_GetGeometryZLevels(P, top, bottom)` — 读取 Z 层级
`gb_SetMaterialFromCopy(P, topZ, bottomZ)` — 从几何创建材料块

### 字符串与本地化
`gv_CTX(dollar, index, default, varType)` — 从 ReverseNest.txt 读取字符串
`gs_ReadAcamCTX()` / `gs_ReadAeditCTX()` / `gs_ReadAcamNestCTX()` — 读取语言文件
`PSDbl(S)` — 字符串→Double
`PSTol(v, places)` — 公差 Double
`PSStr(v, places)` — Double→字符串

### 其他
`gb_AssignImageToGeometry()` / `gb_RemoveImageFromGeometry()` — 几何位图
`gb_HasSubroutines()` — 是否有子程序
`gb_IsFeatureAvailable(sf)` — 实体特征 API 可用
`gb_IsSTLAvailable(oSTL)` — STL 接口可用
`gs_Face(face)` — 工作面名称
`gs_ClampNameFromID(id)` — 夹具名称
`gs_AcamFileType(type)` / `gs_AcamExt(type)` — 文件扩展名
`gb_IsThisActivePost()` — 当前后处理器激活
`g_WipeSolids(bRedraw)` — 删除实体
`g_SetUnsetOpenElement(E, method)` — 打开/闭合元素
`g_InsertAndDrag2DObject(dragType)` — 交互插入拖拽

## 五、通用工具 (modGeneral.bas)

`g_SetAccelerators(Frm, fontSize)` — 快捷键+区域字体
`g_SetCaption(Obj)` — 提取 & 快捷键
`g_SetProperFont(Obj, fontSize)` — 系统语言字体
`gs_GetRegionalSetting(setting)` — 区域设置
`g_Eval(tb, ...)` / `g_EvalInt(tb, ...)` — 数学表达式求值
`gb_IsValOK(sVal, decimal, negative)` — 数字验证
`gb_CheckAllText(Container)` — 文本框非空验证
`gcol_DelimitedStringToCollection(str, delim)` — 分隔字符串→Collection
`gv_Split(str, delim, bBase1)` — 自定义 Split
`gs_NoComma(sVal)` — 逗号→点(德语区域)
`gs_ReplaceSpaces(sVal, sChr)` — 空格替换
`gs_RemoveIllegalChars(sVal, ...)` — 移除非法字符
`gs_RemoveNullChars(sVal)` — 移除 Chr(0)
`gs_DateToString(dt)` — YYYYMMDD
`gs_TruncateText(text, width)` — 中间截断+...
`gs_StripCR` / `gs_StripLF` / `gs_StripCRLF` — 去换行
`gs_GUID()` — 生成 GUID
`gl_BinaryValue(num, binary)` — 2^(n-1) 位掩码
`gl_TranslateColor(clr, hPal)` — OLE_COLOR→RGB
`g_Repaint(Frm)` — 强制重绘
`gl_FrmHwnd(Frm)` — 窗体句柄
`g_Help(chm, index)` — HTML Help
`g_DebugNote(s)` — 调试输出
`g_UnLoadAllForms(bEnd)` — 卸载所有窗体
`g_EnableDisableControls(container, enable)` — 控件启用/禁用
`gb_IsInCollection(C, V)` — Collection 键存在
`g_GetArrayBounds(vArr, lL, lU)` — 数组边界
`gs_NoZeros(sVal)` — 去尾零
`gs_RemovePointFromZero(sVal)` — 0.0→0

## 六、注册表工具 (modRegistry.bas)

`gs_ReadRegKey(keyPath, subKey, root, default)` — 读取
`gb_WriteRegKey(type, keyPath, subKey, value, root)` — 写入
`gs_EnumerateRegKeys(root, keyPath)` / `gs_EnumerateRegKeyValues()` — 枚举
`gb_DeleteRegKey()` / `gb_DeleteRegKeyValue()` — 删除
`gb_ExportRegKey()` / `gb_ImportRegKey()` — 导出/导入

**路径:** HKCU\Software\Planit\Alphacam\ReverseNest\

## 七、文件/路径工具 (modFilesFolders.bas)

`gs_ThisDir()` — 插件所在目录
`gs_ThisFile()` — .amb 文件完整路径
`gs_ReadFileContents(file)` — 读取文本文件
`gs_ParseFileName()` / `gs_ParseDirName()` — 路径解析
`gs_ParseFileExtension()` / `gs_StripFileExtension()` / `gs_ReplaceFileExtension()` — 扩展名操作
`gs_EnsureBackslash()` / `gs_StripLeadingBackslash()` — 斜杠处理
`gs_GetLocalAppDataDir()` / `gs_GetCommonAppDataDir()` — AppData 目录
`gs_GetSpecialFolder(folder, bCreate)` — Windows 特殊文件夹
`gs_GetDir(title, rootDir, ...)` — 文件夹浏览对话框
`gs_MacroDir(macroName)` — VBA 宏目录
`gs_AppDir()` — AlphaCAM 程序目录
`gs_UniqueFileName(file)` — 不重复文件名
`gs_FileSize(file)` — 可读文件大小
`gb_ProjectExists(name, fileName)` — VBA 项目是否加载

## 八、全局声明 (modGlobal.bas)

DEF_APP_TITLE = "Reverse Side Nesting"
DEF_VERSION = "1.2"
DEF_MACRO_NAME = "ReverseNest"
AlphaIntersectPoint 枚举(PARALLEL/NONE/LINE_1/LINE_2/BOTH_LINES)
POINT_XYZ / WP_XYZ / LINE_XYZ / ARC_DETAILS 几何类型
LicomUKsab_* / LicomUKja_* / LicomUKjba_* Nest 属性前缀

## 九、语言文件键值 (ReverseNest.txt)

$1 = "Reverse-Side Nesting"(插件名)
$2 = 插件描述
$10 = 对话框标题: Reverse Side Nesting / Sheet Ordering / Flip Sheet Around
$20 = 选项: By Side / By Sheet / Minimise Tool Changes / X Axis / Y Axis / Include Geometries
$30 = 消息: 无排版 / 无刀具路径 / (reverse) / 打开失败 / 包含工作平面
