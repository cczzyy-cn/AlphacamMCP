' ==============================================================================
' CCC功能 - modMirrorPath 路径自身镜像
' ==============================================================================
' 功能：对用户交互选择的刀具路径，以路径自身的中心线为轴做镜像。
'   - 绕X轴镜像（上下翻转）：以路径 Y 方向中点水平线为轴
'   - 绕Y轴镜像（左右翻转）：以路径 X 方向中点垂直线为轴
'
' 交互流程（参考 modTrim 依边界裁剪）：
'   1. 点击菜单 -> drw.UserSelectMultiGeos 让用户选择刀具路径
'   2. 用户框选/点选后右键完成（或 ESC）
'   3. 弹窗 frmMirrorPath 选择镜像轴 X/Y
'   4. 点击确定 -> 执行镜像
' ==============================================================================
Option Explicit
Option Private Module

' 模块级变量：存储用户交互选择的路径（供 frmMirrorPath.cmdOK_Click 读取）
Private m_selectedPaths As Collection


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
    drw.SetGeosSelected False
    If Not drw.UserSelectMultiGeos("【路径自身镜像】请选择要镜像的刀具路径（框选或点选）", 0) Then
        Exit Sub  ' 用户取消选择
    End If
    
    ' --- 读取选中的路径（仅刀具路径） ---
    Set m_selectedPaths = New Collection
    Dim g As Path
    For Each g In drw.Geometries
        If g.Selected And g.IsToolPath Then
            m_selectedPaths.Add g
        End If
    Next g
    drw.SetGeosSelected False  ' 清除选中状态
    
    ' --- 检查是否选中了有效路径 ---
    If m_selectedPaths.Count = 0 Then
        MsgBox "没有选中任何刀具路径！" & vbCrLf & vbCrLf & _
               "请确保选择的是刀具路径（而非几何图形）。", vbExclamation, "路径自身镜像"
        Exit Sub
    End If
    
    ' --- 弹出对话框选择镜像轴（模态，选择完成后自动关闭） ---
    frmMirrorPath.Show
End Sub


' ==============================================================================
' 返回当前选中的路径集合（供 frmMirrorPath 对话框读取）
' ==============================================================================
Public Function GetSelectedPaths() As Collection
    Set GetSelectedPaths = m_selectedPaths
End Function


' ==============================================================================
' Public Sub DoMirrorPath(ByVal mirrorX As Boolean)
' ==============================================================================
' 执行路径自身镜像。被 frmMirrorPath.cmdOK_Click 调用。
' mirrorX: True=绕X轴（水平线，上下翻转），False=绕Y轴（垂直线，左右翻转）
' ==============================================================================
Public Sub DoMirrorPath(ByVal mirrorX As Boolean)
    On Error GoTo ErrHandler
    
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    
    If m_selectedPaths Is Nothing Then Exit Sub
    If m_selectedPaths.Count = 0 Then Exit Sub
    
    Dim count As Long: count = 0
    Dim tp As Path
    Dim midX As Double, midY As Double
    
    App.SetUndoCommandName "路径自身镜像"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    
    For Each tp In m_selectedPaths
        If Not (tp Is Nothing) Then
            ' 计算路径自身的中心线
            midX = (tp.MinXL + tp.MaxXL) / 2
            midY = (tp.MinYL + tp.MaxYL) / 2
            
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
    Next tp
    
    drw.ScreenUpdating = True
    drw.Redraw
    
    MsgBox "镜像完成！" & vbCrLf & vbCrLf & _
           "已镜像: " & count & " 条刀具路径" & vbCrLf & _
           "镜像轴: " & IIf(mirrorX, "绕X轴（上下翻转）", "绕Y轴（左右翻转）"), _
           vbInformation, "路径自身镜像"
    
    ' 清空模块级变量
    Set m_selectedPaths = Nothing
    Exit Sub
    
ErrHandler:
    drw.ScreenUpdating = True
    Set m_selectedPaths = Nothing
    MsgBox "镜像出错: " & Err.Description, vbCritical, "路径自身镜像"
End Sub
