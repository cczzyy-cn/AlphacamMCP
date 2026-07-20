' ==============================================================================
' CCC功能 - modMirrorPath 路径自身镜像
' ==============================================================================
' 功能：对用户交互选择的刀具路径，以路径自身的中心线为轴做镜像。
'   - 绕X轴镜像（上下翻转）：以路径 Y 方向中点水平线为轴
'   - 绕Y轴镜像（左右翻转）：以路径 X 方向中点垂直线为轴
'
' 交互流程（参考 Drawing.UserSelectMultiToolPaths 文档）：
'   1. 点击菜单 -> drw.UserSelectMultiToolPaths 让用户选择刀具路径
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
    
    ' --- 交互选择：让用户在图纸中框选/点选刀具路径（UserSelectMultiToolPaths） ---
    If Not drw.UserSelectMultiToolPaths("【路径自身镜像】请选择要镜像的刀具路径（框选或点选）", 0) Then
        Exit Sub  ' 用户取消选择
    End If
    
    ' --- 读取选中的刀具路径 ---
    Set m_selectedPaths = New Collection
    Dim tpSel As Path
    For Each tpSel In drw.ToolPaths
        If tpSel.Selected Then
            m_selectedPaths.Add tpSel
            tpSel.Selected = False  ' 清除选中状态（按文档示例）
        End If
    Next tpSel
    
    ' --- 检查是否选中了有效路径 ---
    If m_selectedPaths.Count = 0 Then
        MsgBox "没有选中任何刀具路径！", vbExclamation, "路径自身镜像"
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
    Dim elem As Element
    Dim tpElems As Elements
    
    App.SetUndoCommandName "路径自身镜像"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    
    For Each tp In m_selectedPaths
        If Not (tp Is Nothing) Then
            ' --- 读取所有元素端点，计算几何中心（平均值） ---
            Dim sumX As Double: sumX = 0
            Dim sumY As Double: sumY = 0
            Dim ptCount As Long: ptCount = 0
            
            Set tpElems = tp.Elements
            If tpElems Is Nothing Then GoTo nextTP
            For Each elem In tpElems
                If Not (elem Is Nothing) Then
                    sumX = sumX + elem.StartXL + elem.EndXL
                    sumY = sumY + elem.StartYL + elem.EndYL
                    ptCount = ptCount + 2
                End If
            Next elem
            If ptCount = 0 Then GoTo nextTP
            
            midX = sumX / ptCount
            midY = sumY / ptCount
            
            ' --- 直接在原路径上执行 MoveL+MirrorL+MoveL ---
            ' 用中间变量传递负值（VBA 语法限制）
            Dim nx As Double: nx = -midX
            Dim ny As Double: ny = -midY
            tp.MoveL nx, ny
            
            If mirrorX Then
                tp.MirrorL -10000, 0, 10000, 0
            Else
                tp.MirrorL 0, -10000, 0, 10000
            End If
            
            tp.MoveL midX, midY
            count = count + 1
        End If
nextTP:
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
