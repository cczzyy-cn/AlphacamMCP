' ==============================================================================
' CCC功能 合集 — UserForm 版
' ==============================================================================

' === 功能1：依边界裁剪 ===
Sub 依边界裁剪()
    Dim drw As Drawing
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        MsgBox "没有活动图纸！", vbExclamation, "依边界裁剪"
        Exit Sub
    End If
    
    App.SetUndoCommandName "依边界裁剪"
    App.SetUndoPoint
    
    Dim boundary As Path
    Set boundary = drw.UserSelectOneGeo("【依边界裁剪】请选择边界几何图形（封闭图形）")
    If boundary Is Nothing Then Exit Sub
    
    If Not boundary.Closed Then
        If MsgBox("边界不是封闭图形，继续吗？", vbYesNo + vbQuestion, "依边界裁剪") = vbNo Then
            Exit Sub
        End If
    End If
    
    drw.SetGeosSelected False
    If Not drw.UserSelectMultiGeos("【依边界裁剪】请选择要裁剪的线段（框选或点选）", 0) Then
        Exit Sub
    End If
    
    drw.ScreenUpdating = False
    Dim selectedGeos As New Collection
    Dim g As Path
    For Each g In drw.Geometries
        If g.Selected And (Not (g Is boundary)) Then
            selectedGeos.Add g
        End If
    Next g
    
    boundary.Selected = True
    Dim totalTrimmed As Long: totalTrimmed = 0
    Dim i As Long
    
    For i = 1 To selectedGeos.Count
        Dim geo As Path
        Set geo = selectedGeos(i)
        If geo.TestIntersectPath(boundary, 0, 0) Then
            Dim firstElem As Element: Set firstElem = geo.GetFirstElem
            Dim lastElem As Element: Set lastElem = geo.GetLastElem
            Dim sx As Double, sy As Double, ex As Double, ey As Double
            sx = firstElem.StartXG: sy = firstElem.StartYG
            ex = lastElem.EndXG: ey = lastElem.EndYG
            If Not boundary.IsPointInside(sx, sy) Then
                geo.TrimWithCuttingGeos sx, sy: totalTrimmed = totalTrimmed + 1
            End If
            If Not boundary.IsPointInside(ex, ey) Then
                geo.TrimWithCuttingGeos ex, ey: totalTrimmed = totalTrimmed + 1
            End If
        End If
    Next i
    
    drw.SetGeosSelected False
    drw.ScreenUpdating = True
    drw.Redraw
    drw.ZoomAll
    MsgBox "裁剪完成！已处理 " & totalTrimmed & " 次裁剪。", vbInformation, "依边界裁剪"
End Sub


' === 功能2：全排版刀具偏移（弹出对话框） ===
Sub 全排版刀具偏移()
    Dim drw As Drawing
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        MsgBox "没有活动图纸！", vbExclamation, "全排版刀具偏移"
        Exit Sub
    End If
    
    If drw.Operations Is Nothing Or drw.Operations.Count = 0 Then
        MsgBox "图纸中没有找到任何操作！", vbExclamation, "全排版刀具偏移"
        Exit Sub
    End If
    
    ' 弹出对话框
    frmToolOffset.Show
    
    ' 用户取消
    If frmToolOffset.Tag = "" Then Exit Sub
    
    ' 解析结果
    Dim parts() As String
    parts = Split(frmToolOffset.Tag, "|")
    Dim selectedTool As String: selectedTool = parts(0)
    Dim xOff As Double: xOff = Val(parts(1))
    Dim yOff As Double: yOff = Val(parts(2))
    Dim zOff As Double: zOff = Val(parts(3))
    
    App.SetUndoCommandName "全排版刀具偏移"
    App.SetUndoPoint
    
    ' 执行偏移
    drw.ScreenUpdating = False
    Dim count As Long: count = 0
    Dim ops As Operations: Set ops = drw.Operations
    Dim i As Long, j As Long
    
    For i = 1 To ops.Count
        Dim op As Operation: Set op = ops(i)
        Dim subs As SubOperations: Set subs = op.SubOperations
        If Not (subs Is Nothing) Then
            For j = 1 To subs.Count
                Dim subop As SubOperation: Set subop = subs(j)
                Dim t As MillTool: Set t = subop.Tool
                If Not (t Is Nothing) Then
                    If t.Name = selectedTool Then
                        Dim tps As Paths: Set tps = subop.ToolPaths
                        If Not (tps Is Nothing) Then
                            Dim m As Long
                            For m = 1 To tps.Count
                                Dim tp As Path: Set tp = tps(m)
                                If Not (tp Is Nothing) Then
                                    tp.MoveG xOff, yOff, zOff
                                    count = count + 1
                                End If
                            Next m
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    
    drw.ScreenUpdating = True
    drw.Redraw
    MsgBox "偏移完成！已处理 " & count & " 条刀具路径。" & vbCrLf & _
           "刀具: " & selectedTool & vbCrLf & _
           "偏移量: X=" & xOff & "  Y=" & yOff & "  Z=" & zOff, _
           vbInformation, "全排版刀具偏移"
End Sub


' === 菜单注册 ===
Function InitAlphacamAddIn(AcamVersion As Long) As Integer
    Dim frm As Frame
    Set frm = App.Frame
    With frm
        Dim barId As Long
        barId = .CreateButtonBar("CCC功能")
        .AddMenuItem3 "依边界裁&剪", "m_依边界裁剪", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "cut.bmp", .LastMenuCommandID
        .AddMenuItem3 "全排版刀具偏&移", "m_全排版刀具偏移", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "offset.bmp", .LastMenuCommandID
    End With
    InitAlphacamAddIn = 0
End Function

Function m_依边界裁剪(): 依边界裁剪: End Function
Function m_全排版刀具偏移(): 全排版刀具偏移: End Function
