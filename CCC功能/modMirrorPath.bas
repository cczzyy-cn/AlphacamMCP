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
    Dim pcopy As Path
    Dim elem As Element
    Dim tpElems As Elements
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim eCount As Long
    
    App.SetUndoCommandName "路径自身镜像"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    
    For Each tp In m_selectedPaths
        If Not (tp Is Nothing) Then
            ' --- 遍历所有元素端点，找出 4 个极值点 ---
            xMin = 1E+20: xMax = -1E+20
            yMin = 1E+20: yMax = -1E+20
            eCount = 0
            
            ' 获取当前路径的元素集合
            Set tpElems = tp.Elements
            If tpElems Is Nothing Then GoTo nextTP
            
            For Each elem In tpElems
                If Not (elem Is Nothing) Then
                    ' 起点
                    If elem.StartXL < xMin Then xMin = elem.StartXL
                    If elem.StartXL > xMax Then xMax = elem.StartXL
                    If elem.StartYL < yMin Then yMin = elem.StartYL
                    If elem.StartYL > yMax Then yMax = elem.StartYL
                    ' 终点
                    If elem.EndXL < xMin Then xMin = elem.EndXL
                    If elem.EndXL > xMax Then xMax = elem.EndXL
                    If elem.EndYL < yMin Then yMin = elem.EndYL
                    If elem.EndYL > yMax Then yMax = elem.EndYL
                    eCount = eCount + 1
                End If
            Next elem
            
            If eCount = 0 Then GoTo nextTP
            
            ' 计算路径自身的中心
            midX = (xMin + xMax) / 2
            midY = (yMin + yMax) / 2
            
            ' --- 先减少屏幕刷新防止闪烁，创建副本 ---
            Set pcopy = tp.CopyTemporary
            If Not (pcopy Is Nothing) Then
                ' 策略：MoveL 将几何中心移到局部原点 → MirrorL 过原点 → MoveL 移回
                ' 这避免了 MirrorL 坐标系与元素坐标系的歧义
                
                ' 1. 将路径几何中心移到局部原点
                pcopy.MoveL (-midX), (-midY), 0
                
                ' 2. 在局部原点执行镜像
                If mirrorX Then
                    ' 绕X轴（水平线 y=0，上下翻转）
                    pcopy.MirrorL -10000, 0, 10000, 0
                Else
                    ' 绕Y轴（垂直线 x=0，左右翻转）
                    pcopy.MirrorL 0, -10000, 0, 10000
                End If
                
                ' 3. 移回原位
                pcopy.MoveL midX, midY, 0
                
                ' 存储副本到图纸，删除原始路径
                pcopy.StoreTemporary
                tp.Delete
                count = count + 1
            End If
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
