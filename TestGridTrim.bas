Sub TestGridTrim()
    Dim drw As Drawing
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    
    ' 清空
    Dim cp As Path
    Set cp = drw.GetFirstGeo
    Do While Not (cp Is Nothing)
        cp.Erase
        Set cp = drw.GetFirstGeo
    Loop
    
    ' 画边界（模拟图形3）
    Dim bnd As Path
    Set bnd = drw.CreateRectangle(50, 50, 350, 300)
    bnd.Fillet 10
    
    ' 画网格线
    Dim cx As Double: cx = 200
    Dim y As Double: y = 50
    Dim lines As New Collection
    Do While y <= 350
        Dim l1 As Path
        Set l1 = drw.Create2DLine(cx, y, 350, y + 150)
        lines.Add l1
        Dim l2 As Path
        Set l2 = drw.Create2DLine(cx, y, 50, y + 150)
        lines.Add l2
        Dim l3 As Path
        Set l3 = drw.Create2DLine(cx, y, 350, y - 150)
        lines.Add l3
        Dim l4 As Path
        Set l4 = drw.Create2DLine(cx, y, 50, y - 150)
        lines.Add l4
        y = y + 30
    Loop
    
    MsgBox "已创建 " & lines.Count & " 条网格线，按确定开始裁剪", vbInformation
    
    ' 裁剪
    drw.SetGeosSelected False
    bnd.Selected = True
    
    Dim i As Long
    For i = 1 To lines.Count
        Dim gl As Path
        Set gl = lines(i)
        
        If gl.TestIntersectPath(bnd, 0, 0) Then
            Dim fe As Element
            Set fe = gl.GetFirstElem
            Dim sx As Double, sy As Double
            sx = fe.StartXG: sy = fe.StartYG
            
            drw.SetGeosSelected False
            bnd.Selected = True
            
            If Not bnd.IsPointInside(sx, sy) Then
                gl.TrimWithCuttingGeos sx, sy
            End If
            
            ' 第二端
            drw.SetGeosSelected False
            bnd.Selected = True
            Dim le As Element
            Set le = gl.GetLastElem
            Dim ex As Double, ey As Double
            ex = le.EndXG: ey = le.EndYG
            If Not bnd.IsPointInside(ex, ey) Then
                gl.TrimWithCuttingGeos ex, ey
            End If
        Else
            gl.Erase
        End If
    Next i
    
    drw.SetGeosSelected False
    drw.ZoomAll
    
    ' 验证
    Dim inside As Long, outside As Long
    Set gl = drw.GetFirstGeo
    Do While Not (gl Is Nothing)
        If Not gl.Closed Then
            Dim mx As Double, my As Double, mz As Double
            Dim elem As Element
            If gl.PointAtDistanceAlongPathG(gl.Length / 2, mx, my, mz, elem) Then
                If bnd.IsPointInside(mx, my) Then
                    inside = inside + 1
                Else
                    outside = outside + 1
                End If
            End If
        End If
        Set gl = gl.GetNext
    Loop
    
    MsgBox "边界内: " & inside & "  边界外: " & outside, vbInformation, "TestGridTrim"
End Sub
