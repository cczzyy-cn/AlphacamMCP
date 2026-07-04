' ==============================================================================
' CCC功能 — modRamp 斜角下刀
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

    drw.ScreenUpdating = False: App.SetUndoCommandName "斜角下刀": App.SetUndoPoint

    Set ni = drw.GetNestInformation
    Set ops = drw.Operations
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
                    toolRadius = 3
                    If Not (mt Is Nothing) Then
                        toolRadius = mt.Diameter / 2: If toolRadius <= 0 Then toolRadius = 3
                    End If
                    rampLength = cutDepth / Tan(rampAngle * DEG2RAD)
                    lengthMult = rampLength / toolRadius
                    If lengthMult < 1 Then lengthMult = 1
                    If lengthMult > 50 Then lengthMult = 50

                    ' 设几何起点（朝向排版中心侧的较长边中点）
                    SetStartPointToSheetCenterSide drw, subop, ni, tp

                    ' 选择刀具，选关联几何，删除旧路径，重建新路径
                    If Not (mt Is Nothing) Then
                        Dim tf As String: tf = mt.FileName
                        If tf <> "" Then App.SelectTool tf
                    End If

                    Dim mdNew As MillData: Set mdNew = App.CreateMillData
                    Dim mdOld As MillData: Set mdOld = subop.GetMillData
                    If Not (mdOld Is Nothing) Then
                        mdNew.SafeRapidLevel = mdOld.SafeRapidLevel
                        mdNew.RapidDownTo = mdOld.RapidDownTo
                        mdNew.FinalDepth = -cutDepth
                        mdNew.SpindleSpeed = 24000
                        mdNew.CutFeed = 9000
                        mdNew.DownFeed = 2000
                    End If

                    ' 设起点后选几何重建路径
                    Dim geos As Paths: Set geos = subop.Geometries
                    If Not (geos Is Nothing) Then
                        If geos.Count > 0 Then
                            Dim rampGeo As Path: Set rampGeo = geos(1)
                            If Not (rampGeo Is Nothing) Then
                                oldTp.Delete
                                rampGeo.Selected = True
                                Dim result As Object: Set result = mdNew.RoughFinish
                                rampGeo.Selected = False
                                If Not (result Is Nothing) Then
                                    If result.Count > 0 Then
                                        Dim newTp As Path: Set newTp = result(1)
                                        ' 进刀线（从起点朝排版中心侧延长）
                                        Dim oldTio As Integer: oldTio = newTp.ToolInOut
                                        If oldTio = acamCENTER Then newTp.ToolInOut = acamOUTSIDE
                                        newTp.SetLeadInOutAuto acamLeadBOTH, acamLeadNONE, _
                                                                lengthMult, 0.5, 45, _
                                                                True, False, -0.5
                                        If oldTio = acamCENTER Then newTp.ToolInOut = acamCENTER
                                    End If
                                End If
                            End If
                        End If
                    End If

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
    MsgBox "斜角下刀处理完成！" & vbCrLf & _
           "匹配: " & totalCount & " 条, 小板件: " & smallCount & " 条, 已应用: " & rampApplied & " 条", vbInformation
    Exit Sub
ErrHandler:
    If Not (drw Is Nothing) Then drw.ScreenUpdating = True: drw.Redraw
    MsgBox "斜角下刀出错：" & Err.Description, vbCritical
End Sub

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
                            scx = (sg.MinXL + sg.MaxXL) / 2: scy = (sg.MinYL + sg.MaxYL) / 2
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

    Dim mx As Double: mx = (geo.MinXL + geo.MaxXL) / 2
    Dim my As Double: my = (geo.MinYL + geo.MaxYL) / 2
    Dim dx As Double: dx = scx - mx: dy = scy - my
    Dim startX As Double, startY As Double, maxLen As Double: maxLen = 0

    If dx < 0 Then
        Dim sl As Double: sl = geo.MaxYL - geo.MinYL
        If sl > maxLen Then maxLen = sl: startX = geo.MinXL: startY = my
    End If
    If dx > 0 Then
        Dim sl As Double: sl = geo.MaxYL - geo.MinYL
        If sl > maxLen Then maxLen = sl: startX = geo.MaxXL: startY = my
    End If
    If dy < 0 Then
        Dim sl As Double: sl = geo.MaxXL - geo.MinXL
        If sl > maxLen Then maxLen = sl: startX = mx: startY = geo.MinYL
    End If
    If dy > 0 Then
        Dim sl As Double: sl = geo.MaxXL - geo.MinXL
        If sl > maxLen Then maxLen = sl: startX = mx: startY = geo.MaxYL
    End If
    If maxLen > 0 Then geo.SetStartPoint startX, startY
End Sub
