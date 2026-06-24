' ==============================================================================
' TrimByBoundary — AlphaCAM 依边界裁剪宏
' 功能：用任意封闭边界（矩形/多边形/不规则形状）裁剪穿过它的线段，
'       自动删除边界外的部分。
'
' 原理：对每根穿越边界的线，取其两个端点（都在边界外），
'       用 TrimWithCuttingGeos(X,Y) 依次裁剪两个外端，保留内部段。
'
' 安装：
'   1. 在 AlphaCAM 中按 Alt+F11 打开 VBA 编辑器
'   2. 插入模块，粘贴此代码
'   3. 关闭编辑器，按 Alt+F8 运行 TrimByBoundary
'
' 使用：
'   运行后依次：
'     1. 点击选择边界（封闭图形，如矩形/多边形/不规则形）
'     2. 框选要裁剪的线段
'     3. 自动完成裁剪，弹出完成提示
' ==============================================================================

Sub TrimByBoundary()
    Dim drw As Drawing
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        MsgBox "没有活动图纸！", vbExclamation, "依边界裁剪"
        Exit Sub
    End If
    
    ' 设置撤销点
    App.SetUndoCommandName "依边界裁剪"
    App.SetUndoPoint
    
    ' === 步骤 1：用户选择边界 ===
    Dim boundary As Path
    Set boundary = drw.UserSelectOneGeo("【依边界裁剪】请选择边界几何图形（封闭图形）")
    If boundary Is Nothing Then Exit Sub
    
    ' 验证边界是封闭的
    If Not boundary.Closed Then
        If MsgBox("边界不是封闭图形，继续吗？", vbYesNo + vbQuestion, "依边界裁剪") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' === 步骤 2：用户选择要裁剪的线段 ===
    drw.SetGeosSelected False
    If Not drw.UserSelectMultiGeos("【依边界裁剪】请选择要裁剪的线段（框选或点选）", 0) Then
        Exit Sub
    End If
    
    ' === 步骤 3：执行裁剪 ===
    drw.ScreenUpdating = False
    
    ' 先收集用户选中的几何
    Dim selectedGeos As New Collection
    Dim g As Path
    For Each g In drw.Geometries
        If g.Selected And (Not (g Is boundary)) Then
            selectedGeos.Add g
        End If
    Next g
    
    ' 选中边界作为裁剪工具
    boundary.Selected = True
    
    Dim totalTrimmed As Long
    totalTrimmed = 0
    
    Dim i As Long
    For i = 1 To selectedGeos.Count
        Dim geo As Path
        Set geo = selectedGeos(i)
        
        ' 检查是否与边界相交
        If geo.TestIntersectPath(boundary, 0, 0) Then
            ' 获取线的第一个端点（通常都在边界外）
            Dim firstElem As Element
            Set firstElem = geo.GetFirstElem
            
            ' 获取线的最后一个端点
            Dim lastElem As Element
            Set lastElem = geo.GetLastElem
            
            ' 取两个端点的全局坐标
            Dim sx As Double, sy As Double
            Dim ex As Double, ey As Double
            sx = firstElem.StartXG
            sy = firstElem.StartYG
            ex = lastElem.EndXG
            ey = lastElem.EndYG
            
            ' 判断端点是否在边界外，是则裁剪
            If Not boundary.IsPointInside(sx, sy) Then
                geo.TrimWithCuttingGeos sx, sy
                totalTrimmed = totalTrimmed + 1
            End If
            
            If Not boundary.IsPointInside(ex, ey) Then
                geo.TrimWithCuttingGeos ex, ey
                totalTrimmed = totalTrimmed + 1
            End If
        End If
    Next i
    
    ' 清理
    drw.SetGeosSelected False
    drw.ScreenUpdating = True
    drw.Redraw
    drw.ZoomAll
    
    MsgBox "裁剪完成！已处理 " & totalTrimmed & " 次裁剪（每根线最多2次）。", vbInformation, "依边界裁剪"
End Sub

' ==============================================================================
' 自动注册菜单（保存为 .arb 放入 StartUp/ 后自动加载）
' 仿照 YQ专用 菜单模式：AddMenuItem3 + CreateButtonBar
' ==============================================================================
Function InitAlphacamAddIn(AcamVersion As Long) As Integer
    Dim frm As Frame
    Set frm = App.Frame
    With frm
        Dim barId As Long
        barId = .CreateButtonBar("CCC功能")
        
        ' 添加到 CCC功能 顶级菜单下
        .AddMenuItem3 "依边界裁&剪", "m_TrimByBoundary", acamMenuNEW, "CCC功能", vbNullString
        ' 添加到工具栏
        .AddButton barId, "CCC功能.bmp", .LastMenuCommandID
    End With
    InitAlphacamAddIn = 0
End Function

' 回调函数（点击菜单时调用）
Function m_TrimByBoundary()
    TrimByBoundary
End Function
