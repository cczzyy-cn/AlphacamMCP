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
    Dim drw As Drawing, ops As Operations
    Dim i As Long, j As Long, k As Long
    Dim op As Operation, subs As SubOperations, subop As SubOperation
    Dim mt As MillTool, tps As Paths, tp As Path
    Dim totalCount As Long, smallCount As Long, rampApplied As Long, skipCount As Long
    Dim tpW As Double, tpH As Double, isMatch As Boolean, spPos As Integer, procName As String, selToolNum As Long

    Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！": Exit Sub
    drw.ScreenUpdating = False: App.SetUndoCommandName "斜角下刀": App.SetUndoPoint
    Set ops = drw.Operations
    If ops Is Nothing Or ops.Count = 0 Then drw.ScreenUpdating = True: MsgBox "图纸中没有加工操作！": Exit Sub

    ' 第一遍：收集需要处理的 (tp, subop, mt, origDepth)
    Dim colTP As New Collection, colSO As New Collection, colMT As New Collection, colDepth As New Collection
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
                        colTP.Add tp: colSO.Add subop: colMT.Add mt: colDepth.Add origDepth
                    End If
                End If
NextTp: Next k
NextSub: Next j
NextOp: Next i

    ' 第二遍：逐个处理
    For k = 1 To colTP.Count
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
        Dim steps As Long: steps = CLng(sloopDist / POINT_STEP)
        If steps < 2 Then steps = 2
        Dim finalDepth As Double: finalDepth = actualDepth
        Dim stepDepth As Double: stepDepth = finalDepth / steps
        Dim rampStartDist As Double: rampStartDist = geoLen - sloopDist
        If rampStartDist < 0 Then rampStartDist = 0

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
            steps = CLng(geoLen / POINT_STEP)
            If steps < 2 Then steps = 2
            stepDepth = finalDepth / steps
        End If

        ' 创建 ManualToolPath 从 Z=0 开始
        Dim mtp As Object: Set mtp = mdNew.ManualToolPath(sx, sy, 0#)

        ' 斜坡段：沿路径逐步下刀
        Dim s As Long, px As Double, py As Double, pelem As Element
        For s = 1 To steps
            Dim d As Double: d = rampStartDist + POINT_STEP * s
            If d > geoLen Then d = geoLen
            If toolGeo.PointAtDistanceAlongPathL(d, px, py, pelem) Then
                mtp.Add3DLine px, py, stepDepth * s
            End If
        Next

        ' 跳转到几何起点继续走完
        Dim startX As Double: startX = toolGeo.GetFirstElem.StartXL
        Dim startY As Double: startY = toolGeo.GetFirstElem.StartYL
        mtp.Add3DLine startX, startY, finalDepth

        Dim elems2 As Elements: Set elems2 = toolGeo.Elements
        If Not (elems2 Is Nothing) Then
            For ei = 1 To elems2.Count
                Set elem = elems2(ei)
                If Not (elem Is Nothing) Then
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
