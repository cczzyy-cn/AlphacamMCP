' ==============================================================================
' CCC功能 — modTrim 依边界裁剪
' ==============================================================================
Option Explicit
Option Private Module

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
