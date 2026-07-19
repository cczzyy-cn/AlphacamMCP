' ==============================================================================
' CCC功能 — modMirrorPath 路径自身镜像
' ==============================================================================
' 功能：对当前选中的刀具路径，以路径自身的中心线为轴做镜像。
'   - 绕X轴镜像（上下翻转）：以路径 Y 方向中点水平线为轴
'   - 绕Y轴镜像（左右翻转）：以路径 X 方向中点垂直线为轴
' ==============================================================================
Option Explicit
Option Private Module


' ==============================================================================
' Sub 路径自身镜像()
' ==============================================================================
' 入口过程。由 AlphaCAM 菜单 "CCC功能 > 路径自身镜像" 触发。
' 弹出对话框让用户选择镜像轴。
' ==============================================================================
Sub 路径自身镜像()
    frmMirrorPath.Show vbModeless
End Sub


' ==============================================================================
' Public Sub DoMirrorPath(ByVal mirrorX As Boolean)
' ==============================================================================
' 执行路径自身镜像。
' mirrorX: True=绕X轴（水平线，上下翻转），False=绕Y轴（垂直线，左右翻转）
' ==============================================================================
Public Sub DoMirrorPath(ByVal mirrorX As Boolean)
    On Error GoTo ErrHandler
    
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        MsgBox "没有活动图纸！", vbExclamation, "路径自身镜像"
        Exit Sub
    End If
    
    If drw.GetToolPathCount = 0 Then
        MsgBox "图纸中没有刀具路径！", vbExclamation, "路径自身镜像"
        Exit Sub
    End If
    
    Dim tp As Path
    Dim count As Long: count = 0
    Dim selectedCount As Long: selectedCount = 0
    Dim midX As Double, midY As Double
    
    ' 遍历所有刀具路径，只处理选中状态为 True 的
    Set tp = drw.GetFirstToolPath
    Do While Not (tp Is Nothing)
        If tp.Selected Then
            selectedCount = selectedCount + 1
            
            ' 计算路径自身的中心线
            midX = (tp.MinXL + tp.MaxXL) / 2
            midY = (tp.MinYL + tp.MaxYL) / 2
            
            ' 设置撤销点
            If count = 0 Then
                App.SetUndoCommandName "路径自身镜像"
                App.SetUndoPoint
                drw.ScreenUpdating = False
            End If
            
            ' 执行镜像（以路径自身中心为轴）
            If mirrorX Then
                ' 绕X轴（水平线，上下翻转）
                tp.MirrorL tp.MinXL, midY, tp.MaxXL, midY
            Else
                ' 绕Y轴（垂直线，左右翻转）
                tp.MirrorL midX, tp.MinYL, midX, tp.MaxYL
            End If
            count = count + 1
        End If
        Set tp = tp.GetNext
    Loop
    
    drw.ScreenUpdating = True
    drw.Redraw
    
    If count > 0 Then
        MsgBox "镜像完成！" & vbCrLf & vbCrLf & _
               "选中路径: " & selectedCount & " 条" & vbCrLf & _
               "已镜像: " & count & " 条" & vbCrLf & _
               "镜像轴: " & IIf(mirrorX, "绕X轴（上下翻转）", "绕Y轴（左右翻转）"), _
               vbInformation, "路径自身镜像"
    Else
        MsgBox "没有选中任何刀具路径。" & vbCrLf & vbCrLf & _
               "请先在图纸中选择要镜像的路径（按住 Ctrl 多选），" & vbCrLf & _
               "然后重新点击菜单。", vbInformation, "路径自身镜像"
    End If
    
    Exit Sub
    
ErrHandler:
    drw.ScreenUpdating = True
    MsgBox "镜像出错: " & Err.Description, vbCritical, "路径自身镜像"
End Sub
