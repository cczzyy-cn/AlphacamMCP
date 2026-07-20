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
    Dim tpElems As Elements
    Dim elem As Element
    Dim midX As Double, midY As Double
    Dim xMinG As Double, xMaxG As Double, yMinG As Double, yMaxG As Double
    Dim mtp As Path
    Dim isFirstP As Boolean
    Dim sx As Double, sy As Double, ex As Double, ey As Double
    Dim acx As Double, acy As Double
    
    App.SetUndoCommandName "路径自身镜像"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    
    For Each tp In m_selectedPaths
        If Not (tp Is Nothing) Then
            ' --- 读取所有元素端点的局部坐标，找出极值（用于计算中心） ---
            xMinG = 1E+20: xMaxG = -1E+20
            yMinG = 1E+20: yMaxG = -1E+20
            
            Set tpElems = tp.Elements
            If tpElems Is Nothing Then GoTo skipTP
            For Each elem In tpElems
                If Not (elem Is Nothing) Then
                    If elem.StartXL < xMinG Then xMinG = elem.StartXL
                    If elem.StartXL > xMaxG Then xMaxG = elem.StartXL
                    If elem.EndXL < xMinG Then xMinG = elem.EndXL
                    If elem.EndXL > xMaxG Then xMaxG = elem.EndXL
                    If elem.StartYL < yMinG Then yMinG = elem.StartYL
                    If elem.StartYL > yMaxG Then yMaxG = elem.StartYL
                    If elem.EndYL < yMinG Then yMinG = elem.EndYL
                    If elem.EndYL > yMaxG Then yMaxG = elem.EndYL
                End If
            Next elem
            If xMaxG <= xMinG Or yMaxG <= yMinG Then GoTo skipTP
            
            midX = (xMinG + xMaxG) / 2
            midY = (yMinG + yMaxG) / 2
            
            ' --- MirrorL 对刀具路径不生效，改用 CreateMultiPointPath 手动重建 ---
            ' 使用局部坐标 (StartXL/EndXL)，与 Add3DLine/Add3DArcPointCenter 一致
            Dim mtp As Path: Set mtp = drw.CreateMultiPointPath
            Dim isFirstP As Boolean: isFirstP = True
            
            For Each elem In tpElems
                If Not (elem Is Nothing) Then
                    Dim sx As Double, sy As Double
                    Dim ex As Double, ey As Double
                    ' 计算镜像后的坐标（相对于局部中心）
                    If mirrorX Then
                        ' 上下翻转：y' = 2*midY - y
                        sx = elem.StartXL: sy = 2# * midY - elem.StartYL
                        ex = elem.EndXL: ey = 2# * midY - elem.EndYL
                    Else
                        ' 左右翻转：x' = 2*midX - x
                        sx = 2# * midX - elem.StartXL: sy = elem.StartYL
                        ex = 2# * midX - elem.EndXL: ey = elem.EndYL
                    End If
                    
                    ' 第一个点：从原点移到镜像后的起点
                    If isFirstP Then
                        mtp.Add3DLine sx, sy, 0#
                        isFirstP = False
                    End If
                    
                    ' 添加镜像后的线段或弧线
                    If elem.IsLine Then
                        mtp.Add3DLine ex, ey, 0#
                    ElseIf elem.IsArc Then
                        ' 弧线：需要镜像圆心（局部坐标）
                        Dim acx As Double, acy As Double
                        If mirrorX Then
                            acx = elem.CenterXL: acy = 2# * midY - elem.CenterYL
                        Else
                            acx = 2# * midX - elem.CenterXL: acy = elem.CenterYL
                        End If
                        ' 镜像后圆弧方向翻转
                        mtp.Add3DArcPointCenter ex, ey, 0#, acx, acy, 0#, Not elem.CW
                    End If
                End If
            Next elem
            
            mtp.Finish
            tp.Delete
            count = count + 1
        End If
skipTP:
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
