Dim g_opsPerSheet As Long

' ==============================================================================
' CCC功能 合集 — 最终版
' ==============================================================================

' ==============================================================================
' 功能1：依边界裁剪
' ==============================================================================
Sub 依边界裁剪()
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！", vbExclamation, "依边界裁剪": Exit Sub
    App.SetUndoCommandName "依边界裁剪"
    App.SetUndoPoint
    Dim boundary As Path: Set boundary = drw.UserSelectOneGeo("【依边界裁剪】请选择边界几何图形（封闭图形）")
    If boundary Is Nothing Then Exit Sub
    If Not boundary.Closed Then If MsgBox("边界不是封闭图形，继续吗？", vbYesNo + vbQuestion, "依边界裁剪") = vbNo Then Exit Sub
    drw.SetGeosSelected False
    If Not drw.UserSelectMultiGeos("【依边界裁剪】请选择要裁剪的线段（框选或点选）", 0) Then Exit Sub
    On Error GoTo ErrHandler1
    drw.ScreenUpdating = False
    Dim selectedGeos As New Collection
    Dim g As Path
    For Each g In drw.Geometries
        If g.Selected And (Not (g Is boundary)) Then selectedGeos.Add g
    Next g
    drw.SetGeosSelected False
    boundary.Selected = True
    Dim totalTrimmed As Long
    Dim i As Long
    For i = 1 To selectedGeos.count
        Dim geo As Path: Set geo = selectedGeos(i)
        If geo.Closed Then GoTo SkipGeo1
        If geo.TestIntersectPath(boundary, 0, 0) Then
            Dim firstElem As Element: Set firstElem = geo.GetFirstElem
            Dim lastElem As Element: Set lastElem = geo.GetLastElem
            If firstElem Is Nothing Or lastElem Is Nothing Then GoTo SkipGeo1
            Dim sx As Double, sy As Double, ex As Double, ey As Double
            sx = firstElem.StartXG: sy = firstElem.StartYG
            ex = lastElem.EndXG: ey = lastElem.EndYG
            If Not boundary.IsPointInside(sx, sy) Then geo.TrimWithCuttingGeos sx, sy: totalTrimmed = totalTrimmed + 1
            If Not boundary.IsPointInside(ex, ey) Then geo.TrimWithCuttingGeos ex, ey: totalTrimmed = totalTrimmed + 1
        End If
SkipGeo1:
    Next i
    drw.SetGeosSelected False
    drw.ScreenUpdating = True
    drw.Redraw: drw.ZoomAll
    MsgBox "裁剪完成！已处理 " & totalTrimmed & " 次裁剪。", vbInformation, "依边界裁剪"
    Exit Sub
ErrHandler1:
    drw.ScreenUpdating = True: drw.SetGeosSelected False: drw.Redraw
    MsgBox "裁剪过程中发生错误：" & Err.Description, vbCritical, "依边界裁剪"
End Sub

' ==============================================================================
' 功能2：全排版刀具偏移
' ==============================================================================
Sub 全排版刀具偏移()
    frmToolOffset.Show vbModeless
End Sub

Public Sub ApplyToolOffset(ByVal selectedTool As String, ByVal xOff As Double, ByVal yOff As Double, ByVal zOff As Double)
    On Error GoTo ErrHandler2
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    App.SetUndoCommandName "全排版刀具偏移"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    Dim count As Long
    Dim ops As Operations: Set ops = drw.Operations
    Dim i As Long, j As Long
    For i = 1 To ops.count
        Dim op As Operation: Set op = ops(i)
        Dim subs As SubOperations: Set subs = op.SubOperations
        If Not (subs Is Nothing) Then
            For j = 1 To subs.count
                Dim subop As SubOperation: Set subop = subs(j)
                Dim t As MillTool: Set t = subop.Tool
                If Not (t Is Nothing) Then
                    If t.Name = selectedTool Then
                        Dim tps As paths: Set tps = subop.ToolPaths
                        If Not (tps Is Nothing) Then
                            Dim m As Long
                            For m = 1 To tps.count
                                Dim tp As Path: Set tp = tps(m)
                                If Not (tp Is Nothing) Then tp.MoveG xOff, yOff, zOff: count = count + 1
                            Next m
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    drw.ScreenUpdating = True: drw.Redraw
    MsgBox "偏移完成！已处理 " & count & " 条刀具路径。" & vbCrLf & "刀具: " & selectedTool & vbCrLf & "偏移量: X=" & xOff & "  Y=" & yOff & "  Z=" & zOff, vbInformation, "全排版刀具偏移"
    Exit Sub
ErrHandler2:
    drw.ScreenUpdating = True
    MsgBox "偏移出错：" & Err.Description, vbCritical
End Sub

' ==============================================================================
' 功能3：排版刀具排序
' ==============================================================================
Sub 排版刀具排序()
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！", vbExclamation, "排版刀具排序": Exit Sub
    If drw.Operations Is Nothing Or drw.Operations.count = 0 Then MsgBox "图纸中没有找到任何操作！", vbExclamation, "排版刀具排序": Exit Sub
    frmToolSort.Show vbModeless
End Sub

' ==============================================================================
' 公共辅助函数
' ==============================================================================
Public Function ScanOperations() As Collection
    Dim result As New Collection
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Set ScanOperations = result: Exit Function
    Dim ops As Operations: Set ops = drw.Operations
    If ops Is Nothing Then Set ScanOperations = result: Exit Function
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, j As Long
    For i = 1 To ops.count
        Dim op As Operation: Set op = ops(i)
        Dim subs As SubOperations: Set subs = op.SubOperations
        If subs Is Nothing Then GoTo NextOp2
        For j = 1 To subs.count
            Dim subop As SubOperation: Set subop = subs(j)
            Dim t As MillTool: Set t = subop.Tool
            If Not (t Is Nothing) Then
                Dim methodName As String: methodName = subop.Name
                Dim spPos As Integer
                spPos = InStr(methodName, "  ")
                If spPos > 0 Then methodName = Left(methodName, spPos - 1) Else: spPos = InStr(methodName, " "): If spPos > 0 Then methodName = Left(methodName, spPos - 1)
                Dim key As String: key = methodName & " T" & CStr(t.Number) & " " & t.Name
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        Next j
NextOp2:
    Next i
    Dim keysArr: keysArr = dict.Keys
    Dim kk As Long
    For kk = 0 To UBound(keysArr)
        result.Add keysArr(kk)
    Next kk
    Set ScanOperations = result
End Function

' ★ 最终版：原始算法 + OpNo清零（仅修复二次排序失效）

Public Sub ApplySortToDrawing(ByRef sortedKeys() As String)
    On Error GoTo ErrHandler3
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    Dim ops As Operations: Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub
    Dim mergeDict As Object: Set mergeDict = CreateObject("Scripting.Dictionary")
    Dim opIdx As Long, subIdx As Long
    For opIdx = 1 To ops.count
        Dim opN As Operation: Set opN = ops(opIdx)
        Dim subsN As SubOperations: Set subsN = opN.SubOperations
        If Not (subsN Is Nothing) Then
            For subIdx = 1 To subsN.count
                Dim subN As SubOperation: Set subN = subsN(subIdx)
                Dim tN As MillTool: Set tN = subN.Tool
                If Not (tN Is Nothing) Then
                    Dim mName As String: mName = subN.Name
                    Dim sp As Integer
                    sp = InStr(mName, "  ")
                    If sp > 0 Then mName = Left(mName, sp - 1) Else: sp = InStr(mName, " "): If sp > 0 Then mName = Left(mName, sp - 1)
                    Dim toolKey As String: toolKey = mName & " T" & CStr(tN.Number) & " " & tN.Name
                    Dim dictKey As String: dictKey = CStr(opIdx) & "|" & toolKey
                    If mergeDict.Exists(dictKey) Then mergeDict(dictKey) = mergeDict(dictKey) & "," & CStr(subIdx) Else: mergeDict.Add dictKey, CStr(subIdx)
                End If
            Next subIdx
        End If
    Next opIdx
    Dim allSubOps As New Collection
    Dim keysArr: keysArr = mergeDict.Keys
    Dim ki As Long
    For ki = 0 To UBound(keysArr)
        allSubOps.Add keysArr(ki) & "|" & mergeDict(keysArr(ki))
    Next ki
    If allSubOps.count = 0 Then Exit Sub
    Dim sheetCount As Long: sheetCount = 0
    Dim gi As Long
    For gi = 1 To drw.Geometries.count
        If drw.Geometries(gi).Closed And drw.Geometries(gi).Sheet Then sheetCount = sheetCount + 1
    Next gi
    Dim totalOps As Long: totalOps = ops.count
    Dim opsPerSheet As Long: opsPerSheet = g_opsPerSheet
    If sheetCount > 0 Then opsPerSheet = (totalOps + sheetCount - 1) \ sheetCount
    If g_opsPerSheet = 0 And opsPerSheet > 0 Then g_opsPerSheet = opsPerSheet
    App.SetUndoCommandName "排版刀具排序"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    Dim baseOpNo As Long: baseOpNo = 1
    Dim sheetStart As Long
    For sheetStart = 1 To totalOps Step opsPerSheet
        Dim sheetEnd As Long: sheetEnd = sheetStart + opsPerSheet - 1
        If sheetEnd > totalOps Then sheetEnd = totalOps
        Dim pos As Long: pos = 0
        Dim sj As Long
        For sj = 0 To UBound(sortedKeys)
            Dim targetKey As String: targetKey = sortedKeys(sj)
            For opIdx = sheetStart To sheetEnd
                Dim jj As Long
                For jj = 1 To allSubOps.count
                    Dim parts() As String: parts = Split(allSubOps(jj), "|")
                    Dim recOp As Long: recOp = CLng(parts(0))
                    Dim subKey As String: subKey = parts(1)
                    If recOp = opIdx And subKey = targetKey Then
                        Dim newOpNo As Long: newOpNo = baseOpNo + pos
                        pos = pos + 1
                        Dim subParts() As String: subParts = Split(parts(2), ",")
                        Dim nn As Long
                        For nn = 0 To UBound(subParts)
                            Dim recSub As Long: recSub = CLng(subParts(nn))
                            Dim tgtOp As Operation: Set tgtOp = ops(recOp)
                            Dim tgtSub As SubOperation: Set tgtSub = tgtOp.SubOperations(recSub)
                            Dim tps As paths: Set tps = tgtSub.ToolPaths
                            If Not (tps Is Nothing) Then
                                Dim mm As Long
                                For mm = 1 To tps.count
                                    Dim tp As Path: Set tp = tps(mm)
                                    If Not (tp Is Nothing) Then tp.OpNo = newOpNo
                                Next mm
                            End If
                        Next nn
                    End If
                Next jj
            Next opIdx
        Next sj
        baseOpNo = baseOpNo + pos
    Next sheetStart
    ops.OrderAll
    drw.ScreenUpdating = True: drw.Redraw
    Exit Sub
ErrHandler3:
    drw.ScreenUpdating = True
    MsgBox "排版刀具排序出错：" & Err.Description, vbCritical
End Sub


' ==============================================================================
' 菜单注册
' ==============================================================================
Function InitAlphacamAddIn(AcamVersion As Long) As Integer
    Dim frm As Frame: Set frm = App.Frame
    With frm
        Dim barId As Long: barId = .CreateButtonBar("CCC功能")
        .AddMenuItem3 "依边界裁&剪", "m_依边界裁剪", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "cut.bmp", .LastMenuCommandID
        .AddMenuItem3 "全排版刀具偏&移", "m_全排版刀具偏移", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "offset.bmp", .LastMenuCommandID
        .AddMenuItem3 "排版刀具排&序", "m_排版刀具排序", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "sort.bmp", .LastMenuCommandID
    End With
    InitAlphacamAddIn = 0
End Function

Function m_依边界裁剪(): 依边界裁剪: End Function
Function m_全排版刀具偏移(): 全排版刀具偏移: End Function
Function m_排版刀具排序(): 排版刀具排序: End Function
