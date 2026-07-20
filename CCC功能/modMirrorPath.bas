' ==============================================================================
' CCC功能 - modMirrorPath 路径自身镜像
' ==============================================================================
' 功能：对用户交互选择的刀具路径，以路径自身的中心线为轴做镜像。
'   - 绕X轴镜像（上下翻转）：以路径 Y 方向中点水平线为轴
'   - 绕Y轴镜像（左右翻转）：以路径 X 方向中点垂直线为轴
'
' 交互流程：
'   1. 点击菜单 -> drw.UserSelectMultiToolPaths 让用户选择刀具路径
'   2. 用户框选/点选后右键完成（或 ESC）
'   3. MsgBox 选择镜像轴 X/Y
'   4. 执行镜像
' ==============================================================================
Option Explicit
Option Private Module


' ==============================================================================
' Sub 路径自身镜像()
' ==============================================================================
' 入口过程。由 AlphaCAM 菜单 "CCC功能 > 路径自身镜像" 触发。
' ==============================================================================
Sub 路径自身镜像()
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        MsgBox "没有活动图纸！", vbExclamation, "路径自身镜像"
        Exit Sub
    End If
    
    If drw.GetToolPathCount = 0 Then
        MsgBox "图纸中没有刀具路径！", vbExclamation, "路径自身镜像"
        Exit Sub
    End If
    
    ' --- 交互选择：让用户在图纸中框选/点选刀具路径 ---
    If Not drw.UserSelectMultiToolPaths("【路径自身镜像】请选择要镜像的刀具路径（框选或点选）", 0) Then
        Exit Sub  ' 用户取消选择
    End If
    
    ' --- 不清除选中状态，DoMirrorPath 直接从图纸中读取 Selected ---
    '   （避免用集合缓存路径引用，MirrorL 后引用可能失效）
    
    ' --- MsgBox 选择镜像轴 ---
    If MsgBox("请选择镜像轴：" & vbCrLf & vbCrLf & _
              "是(Y) = 绕X轴（上下翻转）" & vbCrLf & _
              "否(N) = 绕Y轴（左右翻转）", _
              vbYesNo + vbQuestion, "路径自身镜像") = vbYes Then
        DoMirrorPath True   ' 绕X轴
    Else
        DoMirrorPath False  ' 绕Y轴
    End If
End Sub


' ==============================================================================
' Public Sub DoMirrorPath(ByVal mirrorX As Boolean)
' ==============================================================================
' 执行路径自身镜像。
' 从图纸中实时读取 Selected 状态获取要镜像的路径（不用集合缓存引用）。
' mirrorX: True=绕X轴（水平线，上下翻转），False=绕Y轴（垂直线，左右翻转）
' ==============================================================================
Public Sub DoMirrorPath(ByVal mirrorX As Boolean)
    On Error GoTo ErrHandler
    
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    
    Dim count As Long: count = 0
    Dim tp As Path
    Dim midX As Double, midY As Double
    Dim i As Long, n As Long, j As Long
    Dim nameArr() As String
    Dim minXArr() As Double, maxXArr() As Double
    
    App.SetUndoCommandName "路径自身镜像"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    
    ' 第1遍：记录选中路径的 Name + 包围盒（用于交叉验证）
    n = 0
    Set tp = drw.GetFirstToolPath
    Do While Not (tp Is Nothing)
        If tp.Selected Then
            tp.Selected = False
            n = n + 1
            ReDim Preserve nameArr(1 To n)
            ReDim Preserve minXArr(1 To n)
            ReDim Preserve maxXArr(1 To n)
            nameArr(n) = tp.Name
            minXArr(n) = tp.MinXL
            maxXArr(n) = tp.MaxXL
        End If
        Set tp = tp.GetNext
    Loop
    
    ' 第2遍：按 Name 查找路径，验证包围盒，逐一镜像
    For j = 1 To n
        Set tp = drw.GetFirstToolPath
        Do While Not (tp Is Nothing)
            If tp.Name = nameArr(j) Then
                ' 验证包围盒是否匹配（排除 Name 重复或路径错乱）
                If Abs(tp.MinXL - minXArr(j)) < 0.01 And Abs(tp.MaxXL - maxXArr(j)) < 0.01 Then
                    Exit Do
                End If
            End If
            Set tp = tp.GetNext
        Loop
        If Not (tp Is Nothing) Then
            midX = (tp.MinXL + tp.MaxXL) / 2
            midY = (tp.MinYL + tp.MaxYL) / 2
            
            If mirrorX Then
                tp.MirrorL tp.MinXL, midY, tp.MaxXL, midY
            Else
                tp.MirrorL midX, tp.MinYL, midX, tp.MaxYL
            End If
            count = count + 1
        End If
    Next j
    
    drw.ScreenUpdating = True
    drw.Redraw
    
    If count > 0 Then
        MsgBox "镜像完成！" & vbCrLf & "已镜像: " & count & " 条刀具路径" & vbCrLf & _
               "镜像轴: " & IIf(mirrorX, "绕X轴（上下翻转）", "绕Y轴（左右翻转）"), _
               vbInformation, "路径自身镜像"
    Else
        MsgBox "没有选中任何刀具路径！", vbExclamation, "路径自身镜像"
    End If
    
    Exit Sub
    
ErrHandler:
    drw.ScreenUpdating = True
    MsgBox "镜像出错: " & Err.Description, vbCritical, "路径自身镜像"
End Sub
