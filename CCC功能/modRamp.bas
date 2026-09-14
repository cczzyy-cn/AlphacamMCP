' ==============================================================================
' CCC功能 - modRamp 斜角下刀（v2.1.2）
' ==============================================================================
' 依据: 开料小板件吸附与斜下刀算法分析.md
'       开料小板件防松动算法方案.md（HBT 方案）
'
' 【核心约束 - 改本模块前必读】
'   斜坡必须【锚定在轮廓末端】: 从 (轮廓长度 - sloopDist) 处开始下降,
'   在几何起点(闭合点)到达满深度。这样主切削段切掉的是斜坡楔形的【补集】,
'   其吃刀负载在最后 sloopDist 上线性衰减到 0 —— 而轮廓恰在此处闭合,
'   于是「板件被切离的那一瞬间」正好落在整条刀路负载最小的位置上。
'
'   [禁止] 把斜坡改成「以起点为斜坡起始、向前若干毫米」。
'          那会让负载从 0 渐变到满, 切离反而落在满负载上, 比不做斜坡更糟。
'
'   该性质是纯几何的, 与坡角无关; 坡角只影响【斜坡段自身】的负载。
'   只在【闭合刀路】上成立(开放路径的收尾直线会回切到起点)。
'
' v2.1.2 变更（2026-09-14 第 3 轮实机反馈）
'   [A] 【修复·关键】参考中心算错: drw.GetExtent 的参数顺序是 (X1, Y1, Z1, X2, Y2, Z2),
'       第 3 个是 Z1 而不是 X2。原先写成 GetExtent(x1,y1,x2,y2,0,0) 后,
'       X2 位取回的是 Z1(负的切深)、Y2 位取回的是真正的 X2, 于是"参考中心"
'       = (-10, 800) 这种完全错位的点, 排序退化成【按原顺序倒序】。
'       改为正确参数并在取回后校验 (X2>X1 And Y2>Y1), 不合法则用
'       【候选件联合包围盒中心】(整批料的中心)兜底。见 VBA操作问题记录 §4.9。
'   [B] 同一处错误在 SetGeoStartToSheetSide 里也有一份(决定起点落在板件哪条边), 一并修。
'   说明: 抬刀 Z20 在 v2.1.1 实测【已经生效】—— Finish 会把收尾的
'         Add3DRapid 单列成一条 1 元素刀路(rapid, Z -18 -> 20), 紧跟在轮廓刀路后面。
'         设备视角等价, 不是缺陷; 验证脚本要按【子工序内所有刀路拼接】去看末尾元素。
'
' v2.1.1 变更（2026-09-13 第 2 轮实机反馈）
'   [A] 【顺序】主键改回"从外往内": 按【到本板几何中心的 L1 距离】降序(远的先切),
'       HBT 评分降为次键(距离相同才用它)。v2.1.0 拿 HBT 评分当主键, 实机顺序不对。
'   [B] 【抬刀】每条刀路切完补一条 Add3DRapid 到安全高度 Z = 20(用户要求);
'       并把 MillData.SafeRapidLevel 也设成 20 —— 原来沿用原工序, 没保证是 20。
'   [C] 【刷新加工道次窗口】用 Frame.ProjectBarUpdating = False/True 包住整个修改过程。
'       ACAMAPI 原文: "Set to False to stop the project bar being updated ...
'       Set to True when the macro has finished adding paths, the project bar will
'       then be updated" —— ProjectBar 就是加工道次窗口。原先只恢复了 ScreenUpdating,
'       所以窗口一直没刷新。错误分支也保证恢复。
'
' v2.1.0 变更（2026-09-13, 按用户要求）
'   [A] 【移除】微连接(留连接点)/留皮功能 —— 模块与窗体一并删除。
'   [B] 【新增】操作完成后回填加工道次: 用 MillManualToolPath.Finish 返回的 Paths
'       把新刀路的 OpNo 设回原道次, 最后调用 drw.Operations.OrderAll 重排
'       (ACAMAPI 原文: "OpNo - Operation number of this path.
'        Call Operations.OrderAll if this is changed")。
'   [C] 【修复】幂等标记终于写得上了: 原先写在【被删除的旧】路径上, 标记丢失,
'       导致"同一张图重复执行会重复处理"。现在标记写在 Finish 返回的【新】刀路上。
'
' v2.0 变更（承 v1.x）
'   [1] 起点选择修正: 原 SetGeoStartToSheetSide 用「沿边整边跳一次」求点,
'       数学上会跳出包围盒(靠 SetStartPoint 吸附兜住)。改为直接算目标边中点,
'       目标是【朝向排版中心那一侧、较长边】的中点。
'   [2] 处理顺序改为 HBT 评分升序: score = 归一化(轮廓长度) + 归一化(暴露边数)
'       取代原来的「距排版板中心距离」。含义: 释放后新增泄漏小、且被邻件夹得紧的先处理。
'   [3] 小条范围 = 0 时处理【全部闭合刀路】(大板同样需要斜坡; 原实现只处理小板件)。
'   [4] (v2.1.0 已整体移除) 曾新增可选【微连接(留连接点)】—— 按用户要求删除。
'   [5] 新增可选【小件降速】: 按包围盒面积分档降低 CutFeed, 降低侧向推力。
'   [6] 新增 RampVersion() 供部署闭环的项目编译探针调用(无副作用)。
'   [7] 斜坡长度上限保护(0.8×轮廓长度) + rampStartDist 非负保护。
'   [8] 只处理闭合刀路(开放路径计算原属未定义, 原版会回切到起点), 跳过数会报出。
'   [10] v2.0.2 修复: HBT 的"泄漏周长"与降速的"面积"原来用 tp.Length / tp.MinXL~MaxXL,
'       它们【把快速定位(rapid)段也算进去了】(实测: bbox=[0,0,1005,300] 而
'       GetFeedExtent=[405,0]-[1005,300])。现改为: 周长只累加非 rapid 元素长度,
'       位置与面积一律用 GetFeedExtent 的框(取不到时回落包围盒)。
'       另: 由 Element.Length 逐元素累加, 不再依赖 Path.Length。
'   [9] v2.0.1 修复: 无排版的图纸(单件图/测试图)上 drw.GetNestInformation 会抛
'       "现在图档内无排版", 导致整个功能失败 —— 而 SetGeoStartToSheetSide 里
'       本来就有"退回图纸范围中心"的兜底, 却永远走不到。现改为容错取值。
'       同时给 tp.GetFeedExtent / mt.FileName / mt.Diameter 加了取值保护。
'
' [幂等] v2.1.0 起 CCC_RampDone 写在 Finish 返回的【新】刀路上, 重复执行会跳过已处理的
'   刀路(跳过数在完成提示里报出)。仍建议同一张图只跑一次, 或先用 Ctrl+Z 撤销。
'
' 安全: 执行前 App.SetUndoPoint, 一次 Ctrl+Z 可整体撤销本模块的改动。
' ==============================================================================
Option Explicit

' ---- 原路径标记(幂等: v2.1.0 起真正写在新建的刀路上) ----
Private Const ATT_RAMP_DONE    As String = "CCC_RampDone"

' ---- 几何 / 斜坡 ----
' 注意: 不使用长小数字面量。VBA 保存时会把 0.0174532925199433 改写成
'       1.74532925199433E-02, 于是仓库文本与运行文本永远对不上 ——
'       部署闭环的读回校验会因此误报失败。改为运行时计算。
Private Const POINT_STEP       As Double = 0.5
Private Const RAMP_MAX_FRAC    As Double = 0.8    ' 斜坡长度上限 = 该值 x 轮廓长度
' [v2.1.1] 每条刀路切完抬刀到的安全高度(用户要求 Z20)
Private Const SAFE_Z_UP        As Double = 20

' ---- 小件降速(v2.0): 包围盒面积分档(mm^2) ----
Private Const SLOW_A1          As Double = 200000 ' >= 0.20 m^2  不降速
Private Const SLOW_A2          As Double = 50000  ' >= 0.05 m^2  x0.70
Private Const SLOW_A3          As Double = 10000  ' >= 0.01 m^2  x0.50
                                                  ' <  0.01 m^2  x0.35

' ---- 上次窗体值(跨打开保持) ----
Public g_lastMinSize    As Double
Public g_lastCutDepth   As Double
Public g_lastRampAngle  As Double
Public g_lastMethodTool As String
Public g_lastSlowSmall  As Boolean

' ==============================================================================
' 菜单入口
' ==============================================================================
Sub 斜角下刀()
    frmRamp.Show vbModeless
End Sub

' ==============================================================================
' RampVersion - 无副作用, 供部署闭环的项目编译探针 App.Run 调用
' ==============================================================================
Public Function RampVersion() As String
    RampVersion = "modRamp v2.1.2 (2026-09-14)"
End Function

' ==============================================================================
' ApplyRampEntry - 主入口
' ------------------------------------------------------------------------------
' minSize    : 小条判定阈值; <= 0 表示【处理全部闭合刀路】
' cutDepth   : 切割深度(mm, 正数)
' rampAngle  : 斜坡与水平面夹角(度)
' methodName : 加工方式名(空=不限)
' toolMatch  : 刀具匹配串(空=不限)
' tNum       : 刀具号(0=不用)
' (v2.1.0: doTabs/tabStock 两个参数已按用户要求移除)
' slowSmall  : 是否对小件分档降速, 默认 False
' ==============================================================================
Public Sub ApplyRampEntry(ByVal minSize As Double, _
                          ByVal cutDepth As Double, _
                          ByVal rampAngle As Double, _
                          ByVal methodName As String, _
                          ByVal toolMatch As String, _
                          Optional ByVal tNum As Long = 0, _
                          Optional ByVal slowSmall As Boolean = False)
    On Error GoTo ErrHandler

    Dim drw As Drawing, ops As Operations, ni As NestInformation
    Dim i As Long, j As Long, k As Long
    Dim op As Operation, subs As SubOperations, subop As SubOperation
    Dim mt As MillTool, tps As Paths, tp As Path
    Dim totalCount As Long, partCount As Long, rampApplied As Long, skipCount As Long
    Dim openSkipped As Long, noDepth As Long
    Dim slowApplied As Long, markedCount As Long
    Dim tpW As Double, tpH As Double, isMatch As Boolean, spPos As Integer
    Dim procName As String, selToolNum As Long

    ' 收集用(并行集合)
    Dim colTP As Collection, colSO As Collection, colMT As Collection
    Dim colDepth As Collection, colLen As Collection, colOpNo As Collection
    Dim colDist As Collection
    Dim colBX1 As Collection, colBY1 As Collection, colBX2 As Collection, colBY2 As Collection

    ' HBT 评分
    Dim n As Long
    Dim score() As Double, eArr() As Long, ord() As Long
    Dim dist() As Double
    Dim dcx As Double, dcy As Double, refX As Double, refY As Double, blnRef As Boolean
    Dim minL As Double, maxL As Double, minE As Double, maxE As Double
    Dim lv As Double, nl As Double, ne As Double
    Dim a As Long, b As Long, tmpI As Long

    ' 单件处理
    Dim idx As Long, actualDepth As Double, actualDepthAbs As Double
    Dim finalDepth As Double
    ' [v2.1] 加工道次回填(OpNo) + 新刀路的幂等标记
    Dim origOpNo As Long
    Dim newPaths As Paths, lastTp As Path, q As Long
    Dim mdOld As MillData, mdNew As MillData, mdCheck As MillData
    Dim safeR As Double, rapidD As Double, spindle As Double, cutF As Double, downF As Double
    Dim origDepth As Double, depthOk As Boolean
    Dim toolRadius As Double, sloopDist As Double, geoLen As Double
    Dim rampSteps As Long, rampStartDist As Double
    Dim geoObj As Object, toolGeo As Path, elems As Elements, elems2 As Elements
    Dim ei As Long, elem As Element, elem0 As Element, elem2 As Element
    Dim sx As Double, sy As Double, px As Double, py As Double, pelem As Element
    Dim startX As Double, startY As Double
    Dim s As Long, dd As Double, actDist As Double, zz As Double
    Dim fx1 As Double, fy1 As Double, fx2 As Double, fy2 As Double
    Dim blnTake As Boolean, blnClosed As Boolean, blnExt As Boolean
    Dim gLen As Double, area As Double, fac As Double, slowTxt As String
    Dim boxX1 As Double, boxY1 As Double, boxX2 As Double, boxY2 As Double
    Dim els0 As Elements, el0 As Element, ei0 As Long, rapidSegs As Long
    Dim tf As String
    Dim mtp As Object

    Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！": Exit Sub
    drw.ScreenUpdating = False
    App.SetUndoCommandName "斜角下刀"
    App.SetUndoPoint
    Set ops = drw.Operations
    ' [9] 无排版的图纸上 GetNestInformation 会抛 "现在图档内无排版"。
    '     这是正常情形(单件图/测试图), 不是错误: 置 Nothing 后
    '     SetGeoStartToSheetSide 会退回"图纸范围中心", 功能照常可用。
    Set ni = Nothing
    On Error Resume Next
    Set ni = drw.GetNestInformation
    On Error GoTo ErrHandler
    If ops Is Nothing Or ops.Count = 0 Then
        drw.ScreenUpdating = True
        MsgBox "图纸中没有加工操作！" & vbCrLf & _
               "(本模块按 Operations -> SubOperations 遍历;" & vbCrLf & _
               " 若刀路存在但没有加工操作, 这里找不到任何东西。)", vbExclamation, "斜角下刀"
        Exit Sub
    End If

    ' --------------------------------------------------------------------------
    ' 第一遍: 收集候选
    ' --------------------------------------------------------------------------
    Set colTP = New Collection: Set colSO = New Collection: Set colMT = New Collection
    Set colDepth = New Collection: Set colLen = New Collection
    Set colOpNo = New Collection
    Set colDist = New Collection
    Set colBX1 = New Collection: Set colBY1 = New Collection
    Set colBX2 = New Collection: Set colBY2 = New Collection
    totalCount = 0: partCount = 0: rampApplied = 0: skipCount = 0
    openSkipped = 0: noDepth = 0: slowApplied = 0: markedCount = 0

    For i = 1 To ops.Count
        Set op = ops(i)
        Set subs = op.SubOperations
        If subs Is Nothing Then GoTo NextOp
        For j = 1 To subs.Count
            Set subop = subs(j)
            Set mt = subop.Tool
            If mt Is Nothing Then GoTo NextSub
            procName = subop.Name
            spPos = InStr(procName, "  ")
            If spPos > 0 Then
                procName = Left(procName, spPos - 1)
            Else
                spPos = InStr(procName, " ")
                If spPos > 0 Then procName = Left(procName, spPos - 1)
            End If
            If methodName <> "" And procName <> methodName Then GoTo NextSub
            isMatch = False
            If toolMatch <> "" Then
                If mt.Name = toolMatch Then
                    isMatch = True
                ElseIf InStr(1, mt.Name, toolMatch, vbTextCompare) > 0 Then
                    isMatch = True
                ElseIf InStr(1, toolMatch, mt.Name, vbTextCompare) > 0 Then
                    isMatch = True
                ElseIf CStr(mt.Number) = toolMatch Then
                    isMatch = True
                ElseIf tNum > 0 And mt.Number = tNum Then
                    isMatch = True
                ElseIf Left(toolMatch, 1) = "T" Then
                    selToolNum = Val(Mid(toolMatch, 2))
                    If selToolNum > 0 And mt.Number = selToolNum Then isMatch = True
                End If
            Else
                isMatch = True
            End If
            If Not isMatch Then GoTo NextSub
            Set tps = subop.ToolPaths
            If tps Is Nothing Then GoTo NextSub
            For k = 1 To tps.Count
                Set tp = tps(k)
                If tp Is Nothing Then GoTo NextTp
                totalCount = totalCount + 1
                If tp.Attribute(ATT_RAMP_DONE) <> 0 Then skipCount = skipCount + 1: GoTo NextTp
                On Error Resume Next
                blnExt = False
                blnExt = tp.GetFeedExtent(fx1, fy1, fx2, fy2)
                On Error GoTo ErrHandler
                If blnExt Then
                    tpW = fx2 - fx1: tpH = fy2 - fy1
                    ' [10] 位置/面积一律用 feed extent: rapid 段不计入
                    boxX1 = fx1: boxY1 = fy1: boxX2 = fx2: boxY2 = fy2
                Else
                    tpW = tp.MaxXL - tp.MinXL: tpH = tp.MaxYL - tp.MinYL
                    boxX1 = tp.MinXL: boxY1 = tp.MinYL: boxX2 = tp.MaxXL: boxY2 = tp.MaxYL
                End If
                blnTake = False
                If tpW > 1 And tpH > 1 Then
                    If minSize <= 0 Then
                        blnTake = True
                    ElseIf tpW < minSize Or tpH < minSize Then
                        blnTake = True
                    End If
                End If
                ' 只处理闭合刀路; 取不到 Closed 时按闭合处理(保守, 不误杀)
                If blnTake Then
                    blnClosed = True
                    On Error Resume Next
                    blnClosed = tp.Closed
                    On Error GoTo ErrHandler
                    If Not blnClosed Then
                        blnTake = False
                        openSkipped = openSkipped + 1
                    End If
                End If
                If blnTake Then
                    depthOk = False
                    origDepth = -cutDepth
                    Set mdCheck = tp.GetMillData
                    If Not (mdCheck Is Nothing) Then
                        origDepth = CDbl(mdCheck.FinalDepth)
                        If origDepth < 0 And Abs(origDepth) >= cutDepth Then depthOk = True
                    Else
                        depthOk = True          ' 取不到 MillData 也执行(原行为)
                    End If
                    If depthOk Then
                        partCount = partCount + 1
                        colTP.Add tp: colSO.Add subop: colMT.Add mt: colDepth.Add origDepth
                        colOpNo.Add CLng(tp.OpNo)
                        ' [10] 周长只累加【非 rapid】元素长度; 并顺带统计 rapid 段数
                        gLen = 0
                        rapidSegs = 0
                        Set els0 = tp.Elements
                        If Not (els0 Is Nothing) Then
                            For ei0 = 1 To els0.Count
                                Set el0 = els0(ei0)
                                If Not (el0 Is Nothing) Then
                                    If el0.IsRapid Then
                                        rapidSegs = rapidSegs + 1
                                    Else
                                        gLen = gLen + el0.Length
                                    End If
                                End If
                            Next ei0
                        End If
                        colLen.Add gLen
                        colBX1.Add boxX1: colBY1.Add boxY1
                        colBX2.Add boxX2: colBY2.Add boxY2
                    Else
                        noDepth = noDepth + 1
                    End If
                End If
NextTp:
            Next k
NextSub:
        Next j
NextOp:
    Next i

    If colTP.Count = 0 Then
        drw.ScreenUpdating = True
        drw.Redraw
        MsgBox "没有找到符合条件的刀路！" & vbCrLf & _
               "候选=" & totalCount & ", 已处理过=" & skipCount & _
               ", 开放路径跳过=" & openSkipped & ", 深度不足=" & noDepth & _
               IIf(minSize <= 0, "", vbCrLf & "(小条范围=" & minSize & "; 填 0 可处理全部闭合刀路)"), _
               vbInformation, "斜角下刀"
        Exit Sub
    End If

    ' --------------------------------------------------------------------------
    ' 第二遍: HBT 评分 = 归一化(轮廓长度, 泄漏代价) + 归一化(暴露边数, 位移自由度)
    ' --------------------------------------------------------------------------
    n = colTP.Count
    ReDim score(1 To n)
    ReDim eArr(1 To n)
    minL = 1E+30: maxL = -1E+30: minE = 1E+30: maxE = -1E+30
    For k = 1 To n
        eArr(k) = ExposedSides(k, n, colBX1, colBY1, colBX2, colBY2, colMT)
        lv = CDbl(colLen(k))
        If lv < minL Then minL = lv
        If lv > maxL Then maxL = lv
        If CDbl(eArr(k)) < minE Then minE = CDbl(eArr(k))
        If CDbl(eArr(k)) > maxE Then maxE = CDbl(eArr(k))
    Next k
    For k = 1 To n
        nl = 0: ne = 0
        If maxL > minL Then nl = (CDbl(colLen(k)) - minL) / (maxL - minL)
        If maxE > minE Then ne = (CDbl(eArr(k)) - minE) / (maxE - minE)
        score(k) = nl + ne
    Next k
    ' [v2.1.1][A] 主键 = 到【本板几何中心】的 L1 距离, 降序 => 远的先切 = 从外往内
    '             次键 = HBT 评分(升序): 距离相同时, 泄漏小/被夹紧的先切
    ReDim dist(1 To n)
    blnRef = False
    refX = 0: refY = 0
    For k = 1 To n
        If Not FindSheetCenter(drw, ni, colTP(k), dcx, dcy) Then
            ' 没有排版(或该刀路不在任何板上): 退回图纸范围中心
            If Not blnRef Then
                GetRefCenter drw, ni, colBX1, colBY1, colBX2, colBY2, refX, refY
                blnRef = True
            End If
            dcx = refX: dcy = refY
        End If
        dist(k) = Abs(dcx - (CDbl(colBX1(k)) + CDbl(colBX2(k))) / 2) + _
                  Abs(dcy - (CDbl(colBY1(k)) + CDbl(colBY2(k))) / 2)
    Next k
    ReDim ord(1 To n)
    For k = 1 To n: ord(k) = k: Next k
    For a = 2 To n
        tmpI = ord(a)
        b = a - 1
        Do While b >= 1
            If (dist(ord(b)) < dist(tmpI)) Or _
               (dist(ord(b)) = dist(tmpI) And score(ord(b)) > score(tmpI)) Then
                ord(b + 1) = ord(b)
                b = b - 1
            Else
                Exit Do
            End If
        Loop
        ord(b + 1) = tmpI
    Next a

    ' [v2.1.1][C] 包住整个修改过程: 结束后 ProjectBar(加工道次窗口) 会随 True 刷新
    On Error Resume Next
    Frame.ProjectBarUpdating = False
    On Error GoTo ErrHandler

    ' --------------------------------------------------------------------------
    ' 第三遍: 按【从外往内】(到本板中心距离降序) 逐个重建刀路
    ' --------------------------------------------------------------------------
    For k = 1 To n
        idx = ord(k)
        Set tp = colTP(idx)
        Set subop = colSO(idx)
        Set mt = colMT(idx)
        actualDepth = CDbl(colDepth(idx))
        actualDepthAbs = Abs(actualDepth)
        If actualDepthAbs <= 0 Then GoTo SkipItem
        ' [v2.1] 记下原加工道次, 稍后回填给新建的刀路
        origOpNo = CLng(colOpNo(idx))

        If Not (mt Is Nothing) Then
            tf = ""
            On Error Resume Next
            tf = mt.FileName
            On Error GoTo ErrHandler
            If tf <> "" Then App.SelectTool tf
        End If

        ' 读原工序参数(保留原工艺, 不硬编码覆盖)
        safeR = 0: rapidD = 0: spindle = 0: cutF = 0: downF = 0
        Set mdOld = subop.GetMillData
        If Not (mdOld Is Nothing) Then
            safeR = mdOld.SafeRapidLevel: rapidD = mdOld.RapidDownTo
            spindle = mdOld.SpindleSpeed: cutF = mdOld.CutFeed: downF = mdOld.DownFeed
        End If
        If spindle <= 0 Then spindle = 24000
        If cutF <= 0 Then cutF = 9000
        If downF <= 0 Then downF = 2000

        ' [5] 小件降速(可选)
        slowTxt = ""
        If slowSmall Then
            area = (colBX2(idx) - colBX1(idx)) * (colBY2(idx) - colBY1(idx))
            fac = 1#
            If area < SLOW_A3 Then
                fac = 0.35
            ElseIf area < SLOW_A2 Then
                fac = 0.5
            ElseIf area < SLOW_A1 Then
                fac = 0.7
            End If
            If fac < 1# Then
                cutF = CSng(cutF * fac)
                slowApplied = slowApplied + 1
                slowTxt = " F x" & fac
            End If
        End If

        ' 复制几何(跳过快速移动)
        Set geoObj = Nothing
        Set elems = tp.Elements
        If elems Is Nothing Then GoTo SkipItem
        For ei = 1 To elems.Count
            Set elem = elems(ei)
            If Not (elem Is Nothing) Then
                If Not elem.IsRapid Then
                    If geoObj Is Nothing Then
                        Set geoObj = drw.Create2DGeometry(elem.StartXL, elem.StartYL)
                    End If
                    If elem.IsLine Then
                        geoObj.AddLine elem.EndXL, elem.EndYL
                    ElseIf elem.IsArc Then
                        geoObj.AddArcPointCenter elem.EndXL, elem.EndYL, elem.CenterXL, elem.CenterYL, elem.CW
                    End If
                End If
            End If
        Next ei
        If geoObj Is Nothing Then GoTo SkipItem
        Set toolGeo = geoObj.Finish
        If toolGeo Is Nothing Then GoTo SkipItem
        toolGeo.ToolInOut = acamCENTER

        toolRadius = 3
        If Not (mt Is Nothing) Then
            On Error Resume Next
            toolRadius = mt.Diameter / 2
            On Error GoTo ErrHandler
            If toolRadius <= 0 Then toolRadius = 3
        End If

        ' [核心约束] 斜坡锚定在轮廓末端
        sloopDist = actualDepthAbs / Tan(rampAngle * Deg2Rad())
        If sloopDist <= 0 Then sloopDist = 5
        geoLen = toolGeo.Length
        If geoLen <= 0 Then GoTo SkipItem
        If sloopDist > geoLen * RAMP_MAX_FRAC Then sloopDist = geoLen * RAMP_MAX_FRAC
        rampSteps = CLng(sloopDist / POINT_STEP)
        If rampSteps < 2 Then rampSteps = 2
        finalDepth = actualDepth

        rampStartDist = geoLen - sloopDist
        If rampStartDist < 0 Then rampStartDist = 0        ' [7] 非负保护

        ' [1] 起点 = 朝向排版中心那一侧的较长边中点
        SetGeoStartToSheetSide drw, ni, tp, toolGeo

        ' [4] 连接点窗口
        Set mdNew = App.CreateMillData
        ' [v2.1.1][B] 安全高度固定 20(用户要求); safeR 仍读出来但不采用
        mdNew.SafeRapidLevel = SAFE_Z_UP
        mdNew.RapidDownTo = 10
        mdNew.SpindleSpeed = spindle
        mdNew.CutFeed = cutF
        mdNew.DownFeed = downF
        mdNew.FinalDepth = CDbl(finalDepth)

        If Not toolGeo.PointAtDistanceAlongPathL(rampStartDist, sx, sy, elem0) Then
            Set elem0 = toolGeo.GetFirstElem
            If elem0 Is Nothing Then GoTo SkipItem
            sx = elem0.StartXL
            sy = elem0.StartYL
            rampStartDist = 0
        End If

        ' 从 Z=0(板面) 开始, 不是从深度开始 —— 避免直插
        Set mtp = mdNew.ManualToolPath(sx, sy, 0#)

        ' 斜坡段: 沿路径逐步下刀
        For s = 1 To rampSteps
            dd = rampStartDist + POINT_STEP * s
            If dd > geoLen Then dd = geoLen
            actDist = dd - rampStartDist
            zz = -actualDepthAbs * (actDist / sloopDist)
            If toolGeo.PointAtDistanceAlongPathL(dd, px, py, pelem) Then
                mtp.Add3DLine px, py, zz
            End If
        Next s

        startX = toolGeo.GetFirstElem.StartXL
        startY = toolGeo.GetFirstElem.StartYL
        mtp.Add3DLine startX, startY, finalDepth

            Set elems2 = toolGeo.Elements
            If Not (elems2 Is Nothing) Then
                For ei = 1 To elems2.Count
                    Set elem2 = elems2(ei)
                    If Not (elem2 Is Nothing) Then
                        If elem2.IsLine Then
                            mtp.Add3DLine elem2.EndXL, elem2.EndYL, finalDepth
                        ElseIf elem2.IsArc Then
                            mtp.Add3DArcPointCenter elem2.EndXL, elem2.EndYL, finalDepth, _
                                                     elem2.CenterXL, elem2.CenterYL, elem2.CW
                        End If
                    End If
                Next ei
            End If

        ' [v2.1.1][B] 切完抬刀到安全高度 Z20 (Add3DRapid = 快速移动, 不是进给)
        mtp.Add3DRapid startX, startY, SAFE_Z_UP

        ' [v2.1][B][C] Finish 返回新建的 Paths:
        '   把 OpNo 设回原加工道次 (ACAMAPI: "OpNo - Operation number of this path.
        '   Call Operations.OrderAll if this is changed"), 并把幂等标记打在【新】刀路上 ——
        '   这样重复执行会跳过已处理的刀路, 而不是把已经做过斜坡的再做一遍。
        Set newPaths = Nothing
        Set newPaths = mtp.Finish
        If Not (newPaths Is Nothing) Then
            For q = 1 To newPaths.Count
                On Error Resume Next
                newPaths(q).OpNo = CInt(origOpNo)
                newPaths(q).Attribute(ATT_RAMP_DONE) = 1
                On Error GoTo ErrHandler
                markedCount = markedCount + 1
            Next q
        Else
            ' 兜底: Finish 没返回集合时, 用图纸里最后一条刀路
            Set lastTp = drw.GetLastToolPath
            If Not (lastTp Is Nothing) Then
                On Error Resume Next
                lastTp.OpNo = CInt(origOpNo)
                lastTp.Attribute(ATT_RAMP_DONE) = 1
                On Error GoTo ErrHandler
                markedCount = markedCount + 1
            End If
        End If
        toolGeo.Selected = True
        toolGeo.Delete
        tp.Delete
        rampApplied = rampApplied + 1
SkipItem:
    Next k

    ' [v2.1][B] 改过 OpNo 之后必须重排工序(文档: Operations.OrderAll)
    If markedCount > 0 Then
        On Error Resume Next
        drw.Operations.OrderAll
        On Error GoTo ErrHandler
    End If
    ' [v2.1.1][C] 恢复 ProjectBar -> 加工道次窗口随之更新
    On Error Resume Next
    Frame.ProjectBarUpdating = True
    On Error GoTo ErrHandler
    drw.ScreenUpdating = True
    drw.Redraw
    If rampApplied > 0 Then drw.ZoomAll: DoEvents
    MsgBox "斜角下刀处理完成！" & vbCrLf & _
           "候选: " & totalCount & " 条, 受理: " & partCount & " 条, 已应用: " & rampApplied & " 条" & vbCrLf & _
           "跳过: 已处理过 " & skipCount & " / 开放路径 " & openSkipped & " / 深度不足 " & noDepth & vbCrLf & _
           "道次回填: " & markedCount & " 条刀路已设回原 Op 并重排(ProjectBar 已刷新)" & vbCrLf & _
           "抬刀: 每条刀路结束时快速抬到 Z" & SAFE_Z_UP & vbCrLf & _
           IIf(slowSmall, "小件降速: " & slowApplied & " 条" & vbCrLf, "") & _
           "小条范围 = " & IIf(minSize <= 0, "0(全部闭合刀路)", CStr(minSize)) & vbCrLf & _
           "出错或误操作可用一次 Ctrl+Z 整体撤销。", _
           vbInformation, "斜角下刀"
    Exit Sub
ErrHandler:
    ' 先存 Err 再处理(On Error Resume Next 会清空 Err, 见 VBA操作问题记录 4.4)
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    If Not (drw Is Nothing) Then
        On Error Resume Next
        Frame.ProjectBarUpdating = True
        drw.ScreenUpdating = True
        drw.Redraw
    End If
    MsgBox "斜角下刀出错：" & ed & vbCrLf & _
           "(可用一次 Ctrl+Z 撤销本次改动)", vbCritical, "斜角下刀"
End Sub

' ==============================================================================
' FindSheetCenter - 找【该刀路所在排版板】的几何中心(与 v1.x 同口径: 按 OpNo 匹配)
'   返回 False 表示没排版或该刀路不在任何板上 -> 调用方退回图纸范围中心
'   注意: SetGeoStartToSheetSide 里有一段等价的内联查找(为不动已验证代码而保留副本)
' ==============================================================================
Private Function FindSheetCenter(ByVal drw As Drawing, ByVal ni As NestInformation, _
                                 ByVal oldTp As Path, ByRef cx As Double, ByRef cy As Double) As Boolean
    On Error Resume Next
    Dim sh As NestSheet, ps As Paths, sg As Path, pi As Long
    FindSheetCenter = False
    cx = 0: cy = 0
    If ni Is Nothing Then Exit Function
    For Each sh In ni.Sheets
        Set ps = sh.Paths
        If Not (ps Is Nothing) Then
            For pi = 1 To ps.Count
                If ps(pi).OpNo = oldTp.OpNo Then
                    Set sg = sh.Geometry
                    If Not (sg Is Nothing) Then
                        cx = (sg.MinXL + sg.MaxXL) / 2
                        cy = (sg.MinYL + sg.MaxYL) / 2
                        FindSheetCenter = True
                    End If
                    Exit Function
                End If
            Next pi
        End If
    Next sh
End Function

' ==============================================================================
' GetRefCenter - 图纸范围中心(没有排版时的兜底参考点)
' ==============================================================================
Private Sub GetRefCenter(ByVal drw As Drawing, ByVal ni As NestInformation, _
                         ByVal colBX1 As Collection, ByVal colBY1 As Collection, _
                         ByVal colBX2 As Collection, ByVal colBY2 As Collection, _
                         ByRef cx As Double, ByRef cy As Double)
    On Error Resume Next
    Dim sh As NestSheet, sg As Path
    cx = 0: cy = 0
    If Not (ni Is Nothing) Then
        For Each sh In ni.Sheets
            Set sg = sh.Geometry
            If Not (sg Is Nothing) Then
                cx = (sg.MinXL + sg.MaxXL) / 2
                cy = (sg.MinYL + sg.MaxYL) / 2
                Exit Sub
            End If
        Next sh
    End If
    Dim gx1 As Double, gy1 As Double, gz1 As Double
    Dim gx2 As Double, gy2 As Double, gz2 As Double
    ' [v2.1.2] 参数顺序 (X1, Y1, Z1, X2, Y2, Z2) —— 第 3 个是 Z1, 不是 X2!
    drw.GetExtent gx1, gy1, gz1, gx2, gy2, gz2
    If gx2 > gx1 And gy2 > gy1 Then
        cx = (gx1 + gx2) / 2
        cy = (gy1 + gy2) / 2
        Exit Sub
    End If
    ' 兜底: 候选件联合包围盒中心(= 整批料的中心), 不依赖 GetExtent
    Dim kk As Long, ux1 As Double, uy1 As Double, ux2 As Double, uy2 As Double
    If colBX1 Is Nothing Then Exit Sub
    If colBX1.Count = 0 Then Exit Sub
    ux1 = CDbl(colBX1(1)): uy1 = CDbl(colBY1(1))
    ux2 = CDbl(colBX2(1)): uy2 = CDbl(colBY2(1))
    For kk = 2 To colBX1.Count
        If CDbl(colBX1(kk)) < ux1 Then ux1 = CDbl(colBX1(kk))
        If CDbl(colBX2(kk)) > ux2 Then ux2 = CDbl(colBX2(kk))
        If CDbl(colBY1(kk)) < uy1 Then uy1 = CDbl(colBY1(kk))
        If CDbl(colBY2(kk)) > uy2 Then uy2 = CDbl(colBY2(kk))
    Next kk
    cx = (ux1 + ux2) / 2
    cy = (uy1 + uy2) / 2
End Sub

' ==============================================================================
' ExposedSides - 该件 4 条边中有几条【没有邻件】(位移自由度, 越大越容易被推走)
'   邻件判据: 垂直方向有重叠, 且该方向间隙 <= 刀具直径 + 2mm
' ==============================================================================
Private Function ExposedSides(ByVal i As Long, ByVal n As Long, _
                              ByVal colBX1 As Collection, ByVal colBY1 As Collection, _
                              ByVal colBX2 As Collection, ByVal colBY2 As Collection, _
                              ByVal colMT As Collection) As Long
    On Error Resume Next
    Dim gap As Double: gap = 8
    Dim mt As MillTool
    If i >= 1 And i <= colMT.Count Then
        Set mt = colMT(i)
        If Not (mt Is Nothing) Then
            If mt.Diameter > 0 Then gap = mt.Diameter + 2
        End If
    End If
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    x1 = colBX1(i): y1 = colBY1(i): x2 = colBX2(i): y2 = colBY2(i)
    Dim openL As Boolean: openL = True
    Dim openR As Boolean: openR = True
    Dim openB As Boolean: openB = True
    Dim openT As Boolean: openT = True
    Dim j As Long
    Dim ux1 As Double, uy1 As Double, ux2 As Double, uy2 As Double
    For j = 1 To n
        If j <> i Then
            ux1 = colBX1(j): uy1 = colBY1(j): ux2 = colBX2(j): uy2 = colBY2(j)
            ' 左边有邻件: 邻件右缘落在我左缘附近, 且邻件起始在我左侧
            If ux1 < x1 And ux2 >= x1 - gap Then
                If Not (uy2 < y1 Or uy1 > y2) Then openL = False
            End If
            ' 右边有邻件
            If ux2 > x2 And ux1 <= x2 + gap Then
                If Not (uy2 < y1 Or uy1 > y2) Then openR = False
            End If
            ' 下边有邻件
            If uy1 < y1 And uy2 >= y1 - gap Then
                If Not (ux2 < x1 Or ux1 > x2) Then openB = False
            End If
            ' 上边有邻件
            If uy2 > y2 And uy1 <= y2 + gap Then
                If Not (ux2 < x1 Or ux1 > x2) Then openT = False
            End If
        End If
    Next j
    ExposedSides = 0
    If openL Then ExposedSides = ExposedSides + 1
    If openR Then ExposedSides = ExposedSides + 1
    If openB Then ExposedSides = ExposedSides + 1
    If openT Then ExposedSides = ExposedSides + 1
End Function

' ==============================================================================
' ElemLen - Element 没有 Length 属性, 自己算(直线=勾股; 圆弧=R x 扫角)
' ==============================================================================
Private Function ElemLen(ByVal elem As Element) As Double
    On Error Resume Next
    Dim dx As Double, dy As Double, r As Double
    Dim a0 As Double, a1 As Double, sw As Double
    ElemLen = 0
    If elem Is Nothing Then Exit Function
    If elem.IsLine Then
        dx = elem.EndXL - elem.StartXL
        dy = elem.EndYL - elem.StartYL
        ElemLen = Sqr(dx * dx + dy * dy)
        Exit Function
    End If
    r = Sqr((elem.StartXL - elem.CenterXL) ^ 2 + (elem.StartYL - elem.CenterYL) ^ 2)
    If r <= 0 Then Exit Function
    a0 = Atan2(elem.StartYL - elem.CenterYL, elem.StartXL - elem.CenterXL)
    a1 = Atan2(elem.EndYL - elem.CenterYL, elem.EndXL - elem.CenterXL)
    sw = a1 - a0
    If elem.CW Then
        Do While sw > 0
            sw = sw - 2 * Pi()
        Loop
    Else
        Do While sw < 0
            sw = sw + 2 * Pi()
        Loop
    End If
    ElemLen = r * Abs(sw)
End Function

Private Function Pi() As Double
    Pi = 4 * Atn(1)
End Function

Private Function Deg2Rad() As Double
    Deg2Rad = Pi() / 180
End Function

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0 Then
        Atan2 = Atn(y / x)
    ElseIf x < 0 Then
        If y >= 0 Then
            Atan2 = Atn(y / x) + Pi()
        Else
            Atan2 = Atn(y / x) - Pi()
        End If
    Else
        If y >= 0 Then
            Atan2 = Pi() / 2
        Else
            Atan2 = -Pi() / 2
        End If
    End If
End Function

' ==============================================================================
' SetGeoStartToSheetSide - 起点 = 朝向排版中心那一侧的【较长边】的中点
' ------------------------------------------------------------------------------
' v2.0 修正: v1.x 先取最近角、再沿边整边跳一次, 数学上会跳出包围盒
'            (例: 100x600 竖条先选到 (100,300), edgeLen=600 一跳得 (50,-300)),
'            靠 Path.SetStartPoint 吸附兜住才能工作。现在直接算中点, 不再越界。
' 目的: 起点决定【轮廓在哪条边闭合】, 也就是【负载在哪条边归零、件在哪条边被切离】。
'       放在朝向排版中心、且较长的那条边上 —— 那里周围未切材料最多、支撑最好。
' ==============================================================================
Private Sub SetGeoStartToSheetSide(ByVal drw As Drawing, _
                                   ByVal ni As NestInformation, _
                                   ByVal oldTp As Path, _
                                   ByVal toolGeo As Path)
    On Error Resume Next
    Dim scx As Double, scy As Double, found As Boolean
    Dim sh As NestSheet, pInSh As Paths, sg As Path
    Dim pi As Long
    Dim gx1 As Double, gy1 As Double, gz1 As Double
    Dim gx2 As Double, gy2 As Double, gz2 As Double
    Dim mx As Double, my As Double, w As Double, h As Double
    Dim startX As Double, startY As Double
    found = False
    scx = 0: scy = 0
    If Not (ni Is Nothing) Then
        For Each sh In ni.Sheets
            Set pInSh = sh.Paths
            If Not (pInSh Is Nothing) Then
                For pi = 1 To pInSh.Count
                    If pInSh(pi).OpNo = oldTp.OpNo Then
                        Set sg = sh.Geometry
                        If Not (sg Is Nothing) Then
                            scx = (sg.MinXL + sg.MaxXL) / 2
                            scy = (sg.MinYL + sg.MaxYL) / 2
                            found = True
                        End If
                        Exit For
                    End If
                Next pi
            End If
            If found Then Exit For
        Next sh
    End If
    If Not found Then
        ' [v2.1.2] 同上: 参数顺序 (X1, Y1, Z1, X2, Y2, Z2)
        drw.GetExtent gx1, gy1, gz1, gx2, gy2, gz2
        If gx2 > gx1 And gy2 > gy1 Then
            scx = (gx1 + gx2) / 2
            scy = (gy1 + gy2) / 2
        End If
    End If
    mx = (toolGeo.MinXL + toolGeo.MaxXL) / 2
    my = (toolGeo.MinYL + toolGeo.MaxYL) / 2
    w = toolGeo.MaxXL - toolGeo.MinXL
    h = toolGeo.MaxYL - toolGeo.MinYL
    If w >= h Then
        ' 较长边 = 上/下边 -> 取离排版中心更近的那条, 落在它的中点
        startX = mx
        If Abs(scy - toolGeo.MaxYL) <= Abs(scy - toolGeo.MinYL) Then
            startY = toolGeo.MaxYL
        Else
            startY = toolGeo.MinYL
        End If
    Else
        ' 较长边 = 左/右边
        startY = my
        If Abs(scx - toolGeo.MaxXL) <= Abs(scx - toolGeo.MinXL) Then
            startX = toolGeo.MaxXL
        Else
            startX = toolGeo.MinXL
        End If
    End If
    toolGeo.SetStartPoint startX, startY
End Sub
