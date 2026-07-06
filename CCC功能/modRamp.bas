' ==============================================================================
' CCC功能 — modRamp 斜角下刀
' 严格参照小条先切算法：
'   DrawToolGeo 复制刀路元素为几何 → SetStartPoint → ManualToolPath
'   沿原路径的一条边逐渐降低 Z，角度为与水平面的夹角
' ==============================================================================
Option Explicit

Private Const ATT_RAMP_DONE    As String = "CCC_RampDone"
Private Const DEG2RAD          As Double = 0.0174532925199433
Private Const POINT_STEP       As Double = 0.5

' 上次窗体值（跨打开保持）
Public g_lastMinSize    As Double
Public g_lastCutDepth   As Double
Public g_lastRampAngle  As Double
Public g_lastMethodTool As String
Public g_lastEdgeLen    As Double

Sub 斜角下刀()
    frmRamp.Show vbModeless
End Sub

Public Sub ApplyRampEntry(ByVal minSize As Double, _
                          ByVal cutDepth As Double, _
                          ByVal rampAngle As Double, _
                          ByVal methodName As String, _
                          ByVal toolMatch As String, _
                          Optional ByVal tNum As Long = 0)
    On Error GoTo ErrHandler
    Dim drw As Drawing, ops As Operations, ni As NestInformation
    Dim i As Long, j As Long, k As Long
    Dim op As Operation, subs As SubOperations, subop As SubOperation
    Dim mt As MillTool, tps As Paths, tp As Path
    Dim totalCount As Long, smallCount As Long, rampApplied As Long, skipCount As Long
    Dim tpW As Double, tpH As Double, isMatch As Boolean, spPos As Integer, procName As String, selToolNum As Long

    Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！": Exit Sub
    drw.ScreenUpdating = False: App.SetUndoCommandName "斜角下刀": App.SetUndoPoint
    Set ops = drw.Operations
    Set ni = drw.GetNestInformation
    If ops Is Nothing Or ops.Count = 0 Then drw.ScreenUpdating = True: MsgBox "图纸中没有加工操作！": Exit Sub

    ' 第一遍：收集需要处理的 (tp, subop, mt, origDepth)
    Dim colTP As New Collection, colSO As New Collection, colMT As New Collection, colDepth As New Collection, colDist As New Collection
    totalCount = 0: smallCount = 0: rampApplied = 0: skipCount = 0

    For i = 1 To ops.Count
        Set op = ops(i): Set subs = op.SubOperations
        If subs Is Nothing Then GoTo NextOp
        For j = 1 To subs.Count
            Set subop = subs(j): Set mt = subop.Tool
            If mt Is Nothing Then GoTo NextSub
            procName = subop.Name
            spPos = InStr(procName, "  ")
            If spPos > 0 Then procName = Left(procName, spPos - 1) _
            Else: spPos = InStr(procName, " "): If spPos > 0 Then procName = Left(procName, spPos - 1)
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
                If tp.Attribute(ATT_RAMP_DONE) <> 0 Then skipCount = skipCount + 1: GoTo NextTp
                totalCount = totalCount + 1
                Dim fx1 As Double, fy1 As Double, fx2 As Double, fy2 As Double
                If tp.GetFeedExtent(fx1, fy1, fx2, fy2) Then tpW = fx2 - fx1: tpH = fy2 - fy1 _
                Else: tpW = tp.MaxXL - tp.MinXL: tpH = tp.MaxYL - tp.MinYL
                If tpW > 1 And tpH > 1 And (tpW < minSize Or tpH < minSize) Then
                    smallCount = smallCount + 1
                    ' 读取原路径深度（从刀路自身读，更精确）
                    Dim mdCheck As MillData: Set mdCheck = tp.GetMillData
                    Dim depthOk As Boolean: depthOk = False
                    Dim origDepth As Double: origDepth = -cutDepth
                    If Not (mdCheck Is Nothing) Then
                        origDepth = CDbl(mdCheck.FinalDepth)
                        If origDepth < 0 And Abs(origDepth) >= cutDepth Then depthOk = True
                    Else
                        ' 取不到 MillData 也执行
                        depthOk = True
                    End If
                    If depthOk Then
                        ' 计算零件中心到排版中心的距离（远的先切）
                        Dim scX As Double, scY As Double, sf As Boolean: sf = False
                        If Not (ni Is Nothing) Then
                            Dim sh1 As NestSheet
                            For Each sh1 In ni.Sheets
                                Dim ps1 As Paths: Set ps1 = sh1.Paths
                                If Not (ps1 Is Nothing) Then
                                    Dim pi1 As Long
                                    For pi1 = 1 To ps1.Count
                                        If ps1(pi1).OpNo = tp.OpNo Then
                                            Dim sg1 As Path: Set sg1 = sh1.Geometry
                                            If Not (sg1 Is Nothing) Then
                                                scX = (sg1.MinXL + sg1.MaxXL) / 2
                                                scY = (sg1.MinYL + sg1.MaxYL) / 2
                                                sf = True
                                            End If
                                            Exit For
                                        End If
                                    Next pi1
                                End If
                                If sf Then Exit For
                            Next sh1
                        End If
                        Dim partCx As Double: partCx = (tp.MinXL + tp.MaxXL) / 2
                        Dim partCy As Double: partCy = (tp.MinYL + tp.MaxYL) / 2
                        Dim dist As Double: dist = Abs(scX - partCx) + Abs(scY - partCy)
                        colTP.Add tp: colSO.Add subop: colMT.Add mt: colDepth.Add origDepth: colDist.Add dist
                    End If
                End If
NextTp: Next k
NextSub: Next j
NextOp: Next i

    ' 第二遍：逐个处理（远的先切：按距离倒序）
    For k = colTP.Count To 1 Step -1
        Set tp = colTP(k): Set subop = colSO(k): Set mt = colMT(k)
        Dim actualDepth As Double: actualDepth = CDbl(colDepth(k))
        Dim actualDepthAbs As Double: actualDepthAbs = Abs(actualDepth)

        ' 选刀具
        If Not (mt Is Nothing) Then
            Dim tf As String: tf = mt.FileName
            If tf <> "" Then App.SelectTool tf
        End If

        ' 读取 MillData 参数
        Dim mdOld As MillData: Set mdOld = subop.GetMillData
        Dim safeR As Double, rapidD As Double, spindle As Double, cutF As Double, downF As Double
        If Not (mdOld Is Nothing) Then
            safeR = mdOld.SafeRapidLevel: rapidD = mdOld.RapidDownTo
            spindle = mdOld.SpindleSpeed: cutF = mdOld.CutFeed: downF = mdOld.DownFeed
        End If
        If spindle <= 0 Then spindle = 24000
        If cutF <= 0 Then cutF = 9000
        If downF <= 0 Then downF = 2000

        ' 从刀路元素复制几何（参照小条先切 DrawToolGeo，跳过快速移动）
        Dim geoObj As Object: Set geoObj = Nothing
        Dim elems As Elements: Set elems = tp.Elements
        If elems Is Nothing Then GoTo SkipItem
        Dim ei As Long
        For ei = 1 To elems.Count
            Dim elem As Element: Set elem = elems(ei)
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
        Dim toolGeo As Path: Set toolGeo = geoObj.Finish
        If toolGeo Is Nothing Then GoTo SkipItem
        toolGeo.ToolInOut = acamCENTER

        ' 计算斜坡参数
        Dim toolRadius As Double: toolRadius = 3
        If Not (mt Is Nothing) Then
            toolRadius = mt.Diameter / 2
            If toolRadius <= 0 Then toolRadius = 3
        End If
        Dim sloopDist As Double: sloopDist = actualDepthAbs / Tan(rampAngle * DEG2RAD)
        If sloopDist <= 0 Then sloopDist = 5
        Dim geoLen As Double: geoLen = toolGeo.Length
        If geoLen <= 0 Then GoTo SkipItem
        If sloopDist > geoLen * 0.8 Then sloopDist = geoLen * 0.8
        Dim rampSteps As Long: rampSteps = CLng(sloopDist / POINT_STEP)
        If rampSteps < 2 Then rampSteps = 2
        Dim finalDepth As Double: finalDepth = actualDepth

        Dim rampStartDist As Double: rampStartDist = geoLen - sloopDist
        If rampStartDist < 0 Then rampStartDist = 0

        ' 设几何起点在朝向排版中心侧较长边的中点（使斜坡落在该边）
        SetGeoStartToSheetSide drw, subop, ni, tp, toolGeo

        ' Z-18路径后退距离 = 近边长度（留连接点）
        Dim backDist As Double: backDist = Abs(g_lastEdgeLen)
        If backDist <= 0 Then backDist = sloopDist
        rampStartDist = geoLen - sloopDist
        If rampStartDist < 0 Then rampStartDist = 0
        ' 清理临时变量

        ' 创建 MillData
        Dim mdNew As MillData: Set mdNew = App.CreateMillData
        mdNew.SafeRapidLevel = safeR: mdNew.RapidDownTo = 10
        mdNew.SpindleSpeed = spindle: mdNew.CutFeed = cutF: mdNew.DownFeed = downF
        mdNew.FinalDepth = CDbl(finalDepth)

        ' 找斜坡起点坐标
        Dim sx As Double, sy As Double, elem0 As Element
        If Not toolGeo.PointAtDistanceAlongPathL(rampStartDist, sx, sy, elem0) Then
            Set elem0 = toolGeo.GetFirstElem
            If elem0 Is Nothing Then GoTo SkipItem
            sx = elem0.StartXL: sy = elem0.StartYL
            rampStartDist = 0
        End If

        ' 创建 ManualToolPath 从 Z=0 开始
        Dim mtp As Object: Set mtp = mdNew.ManualToolPath(sx, sy, 0#)

        ' 斜坡段：沿路径逐步下刀
        Dim s As Long, px As Double, py As Double, pelem As Element
        Dim d As Double, actDist As Double, z As Double
        For s = 1 To rampSteps
            d = rampStartDist + POINT_STEP * s
            If d > geoLen Then d = geoLen
            actDist = d - rampStartDist
            z = -actualDepthAbs * (actDist / sloopDist)
            If toolGeo.PointAtDistanceAlongPathL(d, px, py, pelem) Then
                mtp.Add3DLine px, py, z
            End If
        Next

        ' 跳转到几何起点继续走完（留 backDist 长度不切到底）
        Dim startX As Double: startX = toolGeo.GetFirstElem.StartXL
        Dim startY As Double: startY = toolGeo.GetFirstElem.StartYL
        mtp.Add3DLine startX, startY, finalDepth

        Dim elems2 As Elements: Set elems2 = toolGeo.Elements
        Dim cumDist As Double: cumDist = 0
        If Not (elems2 Is Nothing) Then
            For ei = 1 To elems2.Count
                Set elem = elems2(ei)
                If Not (elem Is Nothing) Then
                    cumDist = cumDist + elem.Length
                    If cumDist >= geoLen - backDist Then Exit For
                    If elem.IsLine Then
                        mtp.Add3DLine elem.EndXL, elem.EndYL, finalDepth
                    ElseIf elem.IsArc Then
                        mtp.Add3DArcPointCenter elem.EndXL, elem.EndYL, finalDepth, _
                                                 elem.CenterXL, elem.CenterYL, elem.CW
                    End If
                End If
            Next ei
        End If

        mtp.Finish
        toolGeo.Selected = True: toolGeo.Delete

        ' 删除旧刀路
        tp.Delete

        tp.Attribute(ATT_RAMP_DONE) = 1
        rampApplied = rampApplied + 1
SkipItem:
    Next k

    drw.ScreenUpdating = True: drw.Redraw
    If rampApplied > 0 Then drw.ZoomAll: DoEvents
    MsgBox "斜角下刀处理完成！" & vbCrLf & _
           "匹配: " & totalCount & " 条, 小板件: " & smallCount & " 条, 已应用: " & rampApplied & " 条", vbInformation
    Exit Sub
ErrHandler:
    If Not (drw Is Nothing) Then drw.ScreenUpdating = True: drw.Redraw
    MsgBox "斜角下刀出错：" & Err.Description, vbCritical
End Sub

' ==============================================================================
' SetGeoStartToSheetSide — 设几何起点在朝向排版中心侧较长边的中点
' ==============================================================================
Private Sub SetGeoStartToSheetSide(ByVal drw As Drawing, _
                                    ByVal subop As SubOperation, _
                                    ByVal ni As NestInformation, _
                                    ByVal oldTp As Path, _
                                    ByVal toolGeo As Path)
    On Error Resume Next
    Dim scx As Double, scy As Double, found As Boolean: found = False
    If Not (ni Is Nothing) Then
        Dim sh As NestSheet
        For Each sh In ni.Sheets
            Dim pInSh As Paths: Set pInSh = sh.Paths
            If Not (pInSh Is Nothing) Then
                Dim pi As Long
                For pi = 1 To pInSh.Count
                    If pInSh(pi).OpNo = oldTp.OpNo Then
                        Dim sg As Path: Set sg = sh.Geometry
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
        Dim gx1 As Double, gy1 As Double, gx2 As Double, gy2 As Double
        drw.GetExtent gx1, gy1, gx2, gy2, 0, 0
        scx = (gx1 + gx2) / 2: scy = (gy1 + gy2) / 2
    End If
    Dim mx As Double: mx = (toolGeo.MinXL + toolGeo.MaxXL) / 2
    Dim my As Double: my = (toolGeo.MinYL + toolGeo.MaxYL) / 2
    Dim startX As Double, startY As Double
    Dim bestDist As Double: bestDist = 1E+30
    Dim d As Double

    d = Abs(scx - toolGeo.MinXL) + Abs(scy - my)
    If d < bestDist Then bestDist = d: startX = toolGeo.MinXL: startY = my
    d = Abs(scx - toolGeo.MaxXL) + Abs(scy - my)
    If d < bestDist Then bestDist = d: startX = toolGeo.MaxXL: startY = my
    d = Abs(scx - mx) + Abs(scy - toolGeo.MinYL)
    If d < bestDist Then bestDist = d: startX = mx: startY = toolGeo.MinYL
    d = Abs(scx - mx) + Abs(scy - toolGeo.MaxYL)
    If d < bestDist Then bestDist = d: startX = mx: startY = toolGeo.MaxYL

    ' 朝排版中心方向偏移整条边长，落到另一条边的中点上
    Dim edgeLen As Double
    If startX = toolGeo.MinXL Or startX = toolGeo.MaxXL Then
        ' 左/右边 → 沿 Y 偏移边长，X 改为零件中心（落到上/下边的中点）
        edgeLen = toolGeo.MaxYL - toolGeo.MinYL
        If scy > my Then startY = startY + edgeLen Else startY = startY - edgeLen
        startX = mx
    Else
        ' 上/下边 → 沿 X 偏移边长，Y 改为零件中心（落到左/右边的中点）
        edgeLen = toolGeo.MaxXL - toolGeo.MinXL
        If scx > mx Then startX = startX + edgeLen Else startX = startX - edgeLen
        startY = my
    End If

    g_lastEdgeLen = edgeLen
    toolGeo.SetStartPoint startX, startY
End Sub
