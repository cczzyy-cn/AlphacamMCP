' ==============================================================================
' CCC功能 — modRamp 斜角下刀
' ==============================================================================
Option Explicit

Private Const ATT_RAMP_DONE    As String = "CCC_RampDone"
Private Const DEG2RAD          As Double = 0.0174532925199433

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
                    ' 深度 >= 阈值才执行
                    Dim mdCk As MillData: Set mdCk = tp.GetMillData
                    Dim depthOk As Boolean: depthOk = False
                    If Not (mdCk Is Nothing) Then
                        If mdCk.FinalDepth < 0 And Abs(mdCk.FinalDepth) >= cutDepth Then depthOk = True
                    Else
                        depthOk = True
                    End If
                    If depthOk Then
                        Dim toolR As Double: toolR = 3
                        If Not (mt Is Nothing) Then
                            toolR = mt.Diameter / 2
                            If toolR <= 0 Then toolR = 3
                        End If
                        Dim rLen As Double: rLen = cutDepth / Tan(rampAngle * DEG2RAD)
                        Dim lMult As Double: lMult = rLen / toolR
                        If lMult < 1 Then lMult = 1
                        If lMult > 50 Then lMult = 50

                        ' 在关联几何上设起点（4边中距排版中心最近的边，加唯一偏移防重叠）
                        SetUniqueStartPoint drw, subop, ni, tp, k

                        ' 临时切 ToolInOut 再设进刀线
                        Dim oldTio As Integer: oldTio = tp.ToolInOut
                        If oldTio = acamCENTER Then tp.ToolInOut = acamOUTSIDE
                        tp.SetLeadInOutAuto acamLeadBOTH, acamLeadNONE, _
                                            lMult, 0.5, 45, True, False, -0.5
                        If oldTio = acamCENTER Then tp.ToolInOut = acamCENTER

                        tp.Attribute(ATT_RAMP_DONE) = 1
                        rampApplied = rampApplied + 1
                    End If
                End If
NextTp:
            Next k
NextSub:
        Next j
NextOp:
    Next i

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
' SetUniqueStartPoint — 设几何起点在距排版中心最近的边，加唯一偏移避免重复
' ==============================================================================
Private Sub SetUniqueStartPoint(ByVal drw As Drawing, _
                                 ByVal subop As SubOperation, _
                                 ByVal ni As NestInformation, _
                                 ByVal oldTp As Path, _
                                 ByVal idx As Long)
    On Error Resume Next
    Dim geos As Paths: Set geos = subop.Geometries
    If geos Is Nothing Then Exit Sub
    If geos.Count = 0 Then Exit Sub
    Dim geo As Path: Set geo = geos(1)
    If geo Is Nothing Then Exit Sub

    ' 求排版中心
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

    ' 4边中点到排版中心距离，选最近的
    Dim mx As Double: mx = (geo.MinXL + geo.MaxXL) / 2
    Dim my As Double: my = (geo.MinYL + geo.MaxYL) / 2
    Dim startX As Double, startY As Double, best As Double: best = 1E+30
    Dim d As Double
    d = Abs(scx - geo.MinXL) + Abs(scy - my)
    If d < best Then best = d: startX = geo.MinXL: startY = my
    d = Abs(scx - geo.MaxXL) + Abs(scy - my)
    If d < best Then best = d: startX = geo.MaxXL: startY = my
    d = Abs(scx - mx) + Abs(scy - geo.MinYL)
    If d < best Then best = d: startX = mx: startY = geo.MinYL
    d = Abs(scx - mx) + Abs(scy - geo.MaxYL)
    If d < best Then best = d: startX = mx: startY = geo.MaxYL

    ' 沿边偏移 idx*5mm，使每个板的起点不同（防重叠）
    Dim offset As Double: offset = (idx Mod 10) * 5
    If startX = geo.MinXL Or startX = geo.MaxXL Then
        ' 左/右边 → Y 方向偏移
        startY = startY + offset
        If startY > geo.MaxYL - 5 Then startY = geo.MinYL + offset
        If startY < geo.MinYL + 5 Then startY = geo.MinYL + 5
    Else
        ' 上/下边 → X 方向偏移
        startX = startX + offset
        If startX > geo.MaxXL - 5 Then startX = geo.MinXL + offset
        If startX < geo.MinXL + 5 Then startX = geo.MinXL + 5
    End If

    geo.SetStartPoint startX, startY
End Sub
