' ==============================================================================
' CCC功能 — modRamp 斜角下刀
' ==============================================================================
' 核心：SetLeadInOutAuto(SlopeIn=True)
'   1. 遍历 Operations → SubOperations → ToolPaths
'   2. GetFeedExtent 判断小板件
'   3. 找归属排版并计算排版中心
'   4. 判断朝向排版中心的两条边，取较长边
'   5. 在该边中点设 SetStartPoint（使进刀口落在此边）
'   6. 中心铣路径临时切到外侧再设进刀线，避免警告
'   7. Overlap=-0.5 留连接点
' ==============================================================================
Option Explicit
Option Private Module

Private Const ATT_RAMP_DONE    As String = "CCC_RampDone"
Private Const DEG2RAD          As Double = 0.0174532925199433

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
    Dim tpW As Double, tpH As Double, toolRadius As Double, rampLength As Double, lengthMult As Double
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
                    toolRadius = 3
                    If Not (mt Is Nothing) Then
                        toolRadius = mt.Diameter / 2
                        If toolRadius <= 0 Then toolRadius = 3
                    End If
                    rampLength = cutDepth / Tan(rampAngle * DEG2RAD)
                    lengthMult = rampLength / toolRadius
                    If lengthMult < 1 Then lengthMult = 1
                    If lengthMult > 50 Then lengthMult = 50

                    ' 在关联几何上设起点（朝向排版中心侧较长边中点）
                    SetStartPointToSheetCenterSide drw, subop, ni, tp

                    ' 临时切 ToolInOut 再设进刀线（中心铣路径不支持进刀线）
                    Dim oldTio As Integer: oldTio = tp.ToolInOut
                    If oldTio = acamCENTER Then tp.ToolInOut = acamOUTSIDE

                    tp.SetLeadInOutAuto acamLeadBOTH, acamLeadNONE, _
                                        lengthMult, 0.5, 45, _
                                        True, False, -0.5

                    If oldTio = acamCENTER Then tp.ToolInOut = acamCENTER

                    tp.Attribute(ATT_RAMP_DONE) = 1
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
' SetStartPointToSheetCenterSide — 找出朝向排版中心的两条边，取较长边中点设起点
' ==============================================================================
Private Sub SetStartPointToSheetCenterSide(ByVal drw As Drawing, _
                                            ByVal subop As SubOperation, _
                                            ByVal ni As NestInformation, _
                                            ByVal oldTp As Path)

    On Error Resume Next

    Dim geos As Paths: Set geos = subop.Geometries
    If geos Is Nothing Then Exit Sub
    If geos.Count = 0 Then Exit Sub
    Dim geo As Path: Set geo = geos(1)
    If geo Is Nothing Then Exit Sub

    ' 求排版中心
    Dim scx As Double, scy As Double
    Dim found As Boolean: found = False

    If Not (ni Is Nothing) Then
        Dim sh As NestSheet
        For Each sh In ni.Sheets
            Dim pathsInSh As Paths: Set pathsInSh = sh.Paths
            If Not (pathsInSh Is Nothing) Then
                Dim pi As Long
                For pi = 1 To pathsInSh.Count
                    If pathsInSh(pi).OpNo = oldTp.OpNo Then
                        Dim sheetGeo As Path: Set sheetGeo = sh.Geometry
                        If Not (sheetGeo Is Nothing) Then
                            scx = (sheetGeo.MinXL + sheetGeo.MaxXL) / 2
                            scy = (sheetGeo.MinYL + sheetGeo.MaxYL) / 2
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
        ' 降级：用图纸全局中心
        Dim gx1 As Double, gy1 As Double, gx2 As Double, gy2 As Double
        drw.GetExtent gx1, gy1, gx2, gy2, 0, 0
        scx = (gx1 + gx2) / 2: scy = (gy1 + gy2) / 2
    End If

    ' 零件包围盒
    Dim geoMidX As Double: geoMidX = (geo.MinXL + geo.MaxXL) / 2
    Dim geoMidY As Double: geoMidY = (geo.MinYL + geo.MaxYL) / 2
    Dim dx As Double: dx = scx - geoMidX
    Dim dy As Double: dy = scy - geoMidY

    ' 判断哪两条边朝向排版中心，取较长边的中点
    Dim startX As Double, startY As Double
    Dim maxLen As Double: maxLen = 0
    Dim sideLen As Double

    ' 左边（dx<0 表示中心在左）
    If dx < 0 Then
        sideLen = geo.MaxYL - geo.MinYL
        If sideLen > maxLen Then maxLen = sideLen: startX = geo.MinXL: startY = geoMidY
    End If
    ' 右边（dx>0 表示中心在右）
    If dx > 0 Then
        sideLen = geo.MaxYL - geo.MinYL
        If sideLen > maxLen Then maxLen = sideLen: startX = geo.MaxXL: startY = geoMidY
    End If
    ' 下边（dy<0 表示中心在下）
    If dy < 0 Then
        sideLen = geo.MaxXL - geo.MinXL
        If sideLen > maxLen Then maxLen = sideLen: startX = geoMidX: startY = geo.MinYL
    End If
    ' 上边（dy>0 表示中心在上）
    If dy > 0 Then
        sideLen = geo.MaxXL - geo.MinXL
        If sideLen > maxLen Then maxLen = sideLen: startX = geoMidX: startY = geo.MaxYL
    End If

    ' 如果找到了合适边，设置起点
    If maxLen > 0 Then geo.SetStartPoint startX, startY
End Sub
