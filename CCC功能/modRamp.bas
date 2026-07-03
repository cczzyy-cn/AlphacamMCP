' ==============================================================================
' CCC功能 — modRamp 斜角下刀
' 严格参照小条先切 CutToolPath.RoughFinish 算法：
'   1. ManualToolPath 在路径末端 sloopDist 处从 Z=0 开始
'   2. 每 0.5mm 一步逐步下刀至全深
'   3. 跳转到几何起点继续走完整个路径
'   4. 删除旧刀路
' ==============================================================================
Option Explicit
Option Private Module

Private Const ATT_RAMP_DONE    As String = "CCC_RampDone"
Private Const DEG2RAD          As Double = 0.0174532925199433
Private Const POINT_STEP       As Double = 0.5

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
    Dim tpW As Double, tpH As Double, toolRadius As Double
    Dim isMatch As Boolean, spPos As Integer, procName As String, selToolNum As Long

    Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！", vbExclamation, "斜角下刀": Exit Sub

    drw.ScreenUpdating = False
    App.SetUndoCommandName "斜角下刀"
    App.SetUndoPoint

    Set ni = drw.GetNestInformation
    Set ops = drw.Operations
    If ops Is Nothing Or ops.Count = 0 Then
        drw.ScreenUpdating = True: MsgBox "图纸中没有加工操作！", vbExclamation, "斜角下刀": Exit Sub
    End If

    ' 第一阶段：扫描匹配的刀路和几何，存入集合
    Dim colTP As New Collection  ' 存 Path
    Dim colSubop As New Collection  ' 存 SubOperation

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
                If tp.GetFeedExtent(fx1, fy1, fx2, fy2) Then
                    tpW = fx2 - fx1: tpH = fy2 - fy1
                Else
                    tpW = tp.MaxXL - tp.MinXL: tpH = tp.MaxYL - tp.MinYL
                End If

                If tpW > 1 And tpH > 1 And (tpW < minSize Or tpH < minSize) Then
                    smallCount = smallCount + 1
                    colTP.Add tp
                    colSubop.Add subop
                End If
NextTp:
            Next k
NextSub:
        Next j
NextOp:
    Next i

    ' 第二阶段：逐个处理（避免循环中删除刀路导致索引错误）
    For k = 1 To colTP.Count
        Set tp = colTP(k)
        Set subop = colSubop(k)
        CreateRampPath drw, tp, subop, cutDepth, rampAngle
        tp.Attribute(ATT_RAMP_DONE) = 1
        rampApplied = rampApplied + 1
    Next k

    drw.ScreenUpdating = True: drw.Redraw
    If rampApplied > 0 Then drw.ZoomAll: DoEvents
    MsgBox "斜角下刀处理完成！" & vbCrLf & vbCrLf & _
           "匹配刀具路径: " & totalCount & " 条" & vbCrLf & _
           "其中小板件: " & smallCount & " 条" & vbCrLf & _
           "已应用斜角下刀: " & rampApplied & " 条" & vbCrLf & _
           "（跳过已处理: " & skipCount & " 条）", vbInformation, "斜角下刀"
    Exit Sub
ErrHandler:
    If Not (drw Is Nothing) Then drw.ScreenUpdating = True: drw.Redraw
    MsgBox "斜角下刀出错：" & Err.Description, vbCritical, "斜角下刀"
End Sub

' ==============================================================================
' CreateRampPath — 严格参照小条先切 CutToolPath.RoughFinish
' ==============================================================================
Private Sub CreateRampPath(ByVal drw As Drawing, ByVal oldTp As Path, _
                            ByVal subop As SubOperation, _
                            ByVal cutDepth As Double, ByVal rampAngle As Double)

    On Error Resume Next

    ' 取关联几何
    Dim geos As Paths: Set geos = subop.Geometries
    If geos Is Nothing Then Exit Sub
    If geos.Count = 0 Then Exit Sub
    Dim geo As Path: Set geo = geos(1)
    If geo Is Nothing Then Exit Sub
    Dim geoLen As Double: geoLen = geo.Length
    If geoLen <= 0 Then Exit Sub

    ' 读取原有 MillData
    Dim mdOld As MillData: Set mdOld = subop.GetMillData
    If mdOld Is Nothing Then Exit Sub

    ' ===== 1. 斜坡参数（参照小条先切） =====
    Dim sloopDist As Double           ' 斜坡水平距离（mm）
    sloopDist = cutDepth / Tan(rampAngle * DEG2RAD)
    If sloopDist <= 0 Then sloopDist = 5
    If sloopDist > geoLen * 0.8 Then sloopDist = geoLen * 0.8

    Dim steps As Long                 ' 步数 = 斜坡距离 / 0.5
    steps = CLng(sloopDist / POINT_STEP)
    If steps < 2 Then steps = 2

    Dim finalDepth As Double          ' 最终深度（负值）
    finalDepth = -cutDepth

    Dim stepDepth As Double           ' 每步下刀量
    stepDepth = finalDepth / steps

    Dim rampStartDist As Double       ' 斜坡起点距离（沿路径从起点算）
    rampStartDist = geoLen - sloopDist
    If rampStartDist < 0 Then rampStartDist = 0

    ' ===== 2. 创建新 MillData（参照小条先切，用原参数） =====
    Dim mdNew As MillData: Set mdNew = App.CreateMillData
    mdNew.SafeRapidLevel = mdOld.SafeRapidLevel
    mdNew.RapidDownTo = mdOld.RapidDownTo
    mdNew.SpindleSpeed = 24000
    mdNew.CutFeed = 9000
    mdNew.DownFeed = 2000
    mdNew.FinalDepth = finalDepth

    Dim elems As Elements: Set elems = geo.Elements
    If elems Is Nothing Then Exit Sub

    ' ===== 3. 找斜坡起点坐标（参照小条先切 line 354-359） =====
    Dim sx As Double, sy As Double
    Dim elem As Element
    Dim ok As Boolean
    ok = geo.PointAtDistanceAlongPathL(rampStartDist, sx, sy, elem)
    If Not ok Then
        ' 找不到则从头开始
        Set elem = geo.GetFirstElem
        If elem Is Nothing Then Exit Sub
        sx = elem.StartXL: sy = elem.StartYL
        rampStartDist = 0
        steps = CLng(geoLen / POINT_STEP)
        If steps < 2 Then steps = 2
        stepDepth = finalDepth / steps
    End If

    ' ===== 4. 创建 ManualToolPath 从 Z=0 开始（参照小条先切 line 361） =====
    Dim mtp As Object
    Set mtp = mdNew.ManualToolPath(sx, sy, 0#)

    ' ===== 5. 斜坡段：沿路径逐步下刀（参照小条先切 line 363-366） =====
    Dim s As Long, px As Double, py As Double
    Dim pelem As Element
    For s = 1 To steps
        Dim d As Double: d = rampStartDist + POINT_STEP * s
        If d > geoLen Then d = geoLen
        If geo.PointAtDistanceAlongPathL(d, px, py, pelem) Then
            mtp.Add3DLine px, py, stepDepth * s
        End If
    Next

    ' ===== 6. 跳转到几何起点继续走完（参照小条先切 line 368-376） =====
    Dim startX As Double: startX = geo.GetFirstElem.StartXL
    Dim startY As Double: startY = geo.GetFirstElem.StartYL
    mtp.Add3DLine startX, startY, finalDepth

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

    ' ===== 7. 完成 =====
    mtp.Finish

    ' 删除旧刀路
    oldTp.Delete

    Set mdNew = Nothing
    Set mtp = Nothing
End Sub
