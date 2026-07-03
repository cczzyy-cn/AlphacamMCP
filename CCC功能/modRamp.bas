' ==============================================================================
' CCC功能 — modRamp 斜角下刀
' ==============================================================================
' 算法（参照小条先切 CutToolPath.RoughFinish）：
'   1. 遍历 Operations → SubOperations → ToolPaths
'   2. GetFeedExtent 判断小板件（忽略快速移动）
'   3. 对小板件：
'      a. 取关联几何 subop.Geometries(1)
'      b. 在几何上找 (PathLength - sloopDist) 处为斜坡起点
'      c. 创建 ManualToolPath：从 Z+10 高度开始，逐步下刀
'      d. 走完整个路径，删除旧路径
' ==============================================================================
Option Explicit
Option Private Module

Private Const ATT_RAMP_DONE    As String = "CCC_RampDone"
Private Const DEG2RAD          As Double = 0.0174532925199433
Private Const POINT_STEP       As Double = 0.5       ' 每步间距 (mm)

' ==============================================================================
' Sub 斜角下刀() — 入口（由 Events.bas 菜单调用）
' ==============================================================================
Sub 斜角下刀()
    frmRamp.Show vbModeless
End Sub

' ==============================================================================
' Public Sub ApplyRampEntry() — 执行斜角下刀算法
' ==============================================================================
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
    Dim totalCount As Long, smallCount As Long, rampApplied As Long
    Dim skipCount As Long
    Dim tpW As Double, tpH As Double
    Dim toolRadius As Double, rampLength As Double, lengthMult As Double
    Dim isMatch As Boolean, spPos As Integer, procName As String
    Dim selToolNum As Long

    Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！", vbExclamation, "斜角下刀": Exit Sub

    drw.ScreenUpdating = False
    App.SetUndoCommandName "斜角下刀"
    App.SetUndoPoint

    Set ops = drw.Operations
    If ops Is Nothing Or ops.Count = 0 Then
        drw.ScreenUpdating = True: MsgBox "图纸中没有加工操作！", vbExclamation, "斜角下刀": Exit Sub
    End If

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
                If mt.Name = toolMatch Then isMatch = True
                ElseIf InStr(1, mt.Name, toolMatch, vbTextCompare) > 0 Then isMatch = True
                ElseIf InStr(1, toolMatch, mt.Name, vbTextCompare) > 0 Then isMatch = True
                ElseIf CStr(mt.Number) = toolMatch Then isMatch = True
                ElseIf tNum > 0 And mt.Number = tNum Then isMatch = True
                ElseIf Left(toolMatch, 1) = "T" Then
                    selToolNum = Val(Mid(toolMatch, 2))
                    If selToolNum > 0 And mt.Number = selToolNum Then isMatch = True
                End If
            Else: isMatch = True
            End If
            If Not isMatch Then GoTo NextSub

            Set tps = subop.ToolPaths
            If tps Is Nothing Then GoTo NextSub

            For k = 1 To tps.Count
                Set tp = tps(k)
                If tp Is Nothing Then GoTo NextTp
                If tp.Attribute(ATT_RAMP_DONE) <> 0 Then skipCount = skipCount + 1: GoTo NextTp

                totalCount = totalCount + 1

                ' === GetFeedExtent 判断小板件 ===
                Dim fx1 As Double, fy1 As Double, fx2 As Double, fy2 As Double
                If tp.GetFeedExtent(fx1, fy1, fx2, fy2) Then
                    tpW = fx2 - fx1: tpH = fy2 - fy1
                Else
                    tpW = tp.MaxXL - tp.MinXL: tpH = tp.MaxYL - tp.MinYL
                End If

                If tpW > 1 And tpH > 1 And (tpW < minSize Or tpH < minSize) Then
                    smallCount = smallCount + 1

                    toolRadius = 3
                    If Not (mt Is Nothing) Then
                        toolRadius = mt.Diameter / 2
                        If toolRadius <= 0 Then toolRadius = 3
                    End If

                    ' === 斜角下刀：ManualToolPath 方式（参照小条先切） ===
                    Call CreateRampPath drw, tp, subop, cutDepth, rampAngle
                    rampApplied = rampApplied + 1
                End If
NextTp:
            Next k
NextSub:
        Next j
NextOp:
    Next i

    drw.ScreenUpdating = True: drw.Redraw
    If rampApplied > 0 Then drw.ZoomAll: DoEvents

    Dim msg As String
    msg = "斜角下刀处理完成！" & vbCrLf & vbCrLf & _
          "匹配刀具路径: " & totalCount & " 条" & vbCrLf & _
          "其中小板件: " & smallCount & " 条" & vbCrLf & _
          "已应用斜角下刀: " & rampApplied & " 条" & vbCrLf & _
          "（跳过已处理: " & skipCount & " 条）" & vbCrLf & vbCrLf & _
          "参数: 小条范围=" & minSize & "mm, 切割深度=" & cutDepth & "mm, 下刀角度=" & rampAngle & ChrW(176)
    If methodName <> "" Then msg = msg & vbCrLf & "加工方式=" & methodName
    MsgBox msg, vbInformation, "斜角下刀"
    Exit Sub

ErrHandler:
    If Not (drw Is Nothing) Then drw.ScreenUpdating = True: drw.Redraw
    MsgBox "斜角下刀出错：" & Err.Description, vbCritical, "斜角下刀"
End Sub

' ==============================================================================
' CreateRampPath — 用 ManualToolPath 创建斜角下刀路径
' 参照小条先切：在路径末端 sloopDist 处从 Z+10 逐步下刀至全深
' ==============================================================================
Private Sub CreateRampPath(ByVal drw As Drawing, _
                           ByVal oldTp As Path, _
                           ByVal subop As SubOperation, _
                           ByVal cutDepth As Double, _
                           ByVal rampAngle As Double)

    On Error GoTo ErrRamp

    ' 取关联几何
    Dim geos As Paths: Set geos = subop.Geometries
    If geos Is Nothing Then GoTo Fallback
    If geos.Count = 0 Then GoTo Fallback
    Dim geo As Path: Set geo = geos(1)
    If geo Is Nothing Then GoTo Fallback

    Dim geoLen As Double: geoLen = geo.Length
    If geoLen <= 0 Then GoTo Fallback

    ' 斜坡参数
    Dim sloopDist As Double            ' 斜坡水平距离
    sloopDist = cutDepth / Tan(rampAngle * DEG2RAD)
    If sloopDist <= 0 Then sloopDist = 5
    If sloopDist >= geoLen * 0.9 Then sloopDist = geoLen * 0.9

    Dim steps As Long: steps = CLng(sloopDist / POINT_STEP)
    If steps < 2 Then steps = 2
    Dim stepZ As Double: stepZ = -(cutDepth + 10) / steps
    Dim rampStart As Double: rampStart = geoLen - sloopDist
    If rampStart < 0 Then rampStart = 0

    ' 读取原有 MillData 参数
    Dim md As MillData: Set md = subop.GetMillData
    If md Is Nothing Then GoTo Fallback

    Dim safeRapid As Double: safeRapid = md.SafeRapidLevel
    Dim rapidDown As Double: rapidDown = md.RapidDownTo
    Dim finalDepth As Double: finalDepth = -cutDepth
    Dim spindle As Double: spindle = md.SpindleSpeed
    Dim cutFeed As Double: cutFeed = md.CutFeed
    Dim downFeed As Double: downFeed = md.DownFeed
    If spindle <= 0 Then spindle = 18000
    If cutFeed <= 0 Then cutFeed = 9000

    ' 删除旧刀路
    oldTp.Delete

    ' 创建新 MillData
    Dim mdNew As MillData: Set mdNew = App.CreateMillData
    mdNew.SafeRapidLevel = safeRapid
    mdNew.RapidDownTo = rapidDown
    mdNew.SpindleSpeed = spindle
    mdNew.CutFeed = cutFeed
    mdNew.DownFeed = downFeed
    mdNew.FinalDepth = finalDepth

    ' 找斜坡起点（在路径上，不超出原路径范围）
    Dim sx As Double, sy As Double
    Dim elem As Element
    Dim ok As Boolean: ok = geo.PointAtDistanceAlongPathL(rampStart, sx, sy, elem)

    ' 创建 ManualToolPath，从 Z+10 高度开始
    Dim mtp As Object
    If ok Then
        Set mtp = mdNew.ManualToolPath(sx, sy, 10#)
    Else
        Set elem = geo.GetFirstElem
        If elem Is Nothing Then GoTo Fallback
        rampStart = 0
        Set mtp = mdNew.ManualToolPath(elem.StartXL, elem.StartYL, 10#)
    End If

    ' 斜坡段：从 Z+10 逐步下刀到 Z=finalDepth
    Dim s As Long, px As Double, py As Double
    Dim pelem As Element
    For s = 1 To steps
        Dim d As Double: d = rampStart + POINT_STEP * s
        If d > geoLen Then d = geoLen
        If geo.PointAtDistanceAlongPathL(d, px, py, pelem) Then
            mtp.Add3DLine px, py, 10# + stepZ * s
        End If
    Next

    ' 从斜坡结束点走完整个路径
    Dim elems As Elements: Set elems = geo.Elements
    If Not (elems Is Nothing) Then
        Dim ei As Long
        For ei = 1 To elems.Count
            Set elem = elems(ei)
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
    Exit Sub

Fallback:
    ' 降级方案：删除旧路径后直接走 RoughFinish
    oldTp.Delete
    geo.Selected = True
    Dim mdFb As MillData: Set mdFb = App.CreateMillData
    If Not (subop.GetMillData Is Nothing) Then
        Dim mdOld As MillData: Set mdOld = subop.GetMillData
        mdFb.SafeRapidLevel = mdOld.SafeRapidLevel
        mdFb.RapidDownTo = mdOld.RapidDownTo
        mdFb.FinalDepth = -cutDepth
        geo.Selected = True
        mdFb.RoughFinish
        geo.Selected = False
    End If
    Exit Sub

ErrRamp:
    Exit Sub
End Sub
