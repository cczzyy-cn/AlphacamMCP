' ==============================================================================
' CCC功能合集 — 依边界裁剪 + 全排版刀具偏移（UserForm 版）
'
' 这是 frmToolOffset 用户窗体的代码文件
' 功能：弹出一个对话框，让用户选择要偏移的刀具以及 X/Y/Z 三个方向的偏移量
' ==============================================================================

' ==============================================================================
' UserForm_Initialize — 窗体加载时自动执行
' 用途：初始化下拉列表（收集图纸中所有可用的刀具），设置默认偏移值为 0
' ==============================================================================
Private Sub UserForm_Initialize()
    ' --- 获取当前活动图纸 ---
    Dim drw As Drawing
    Set drw = App.ActiveDrawing
    ' 没有图纸则直接退出初始化（窗体将显示为空）
    If drw Is Nothing Then Exit Sub
    
    ' --- 获取图纸中的所有加工操作（Operations） ---
    Dim ops As Operations
    Set ops = drw.Operations
    ' 没有操作也直接退出
    If ops Is Nothing Or ops.count = 0 Then Exit Sub
    
    Dim i As Long, j As Long
    
    ' --- 遍历所有操作，收集操作中使用的刀具信息 ---
    ' AlphaCAM 操作层级：Operations → SubOperations → Tool（刀具）
    ' 需要两层循环来访问所有子操作
    For i = 1 To ops.count
        Dim op As Operation
        Set op = ops(i)
        Dim subs As SubOperations
        Set subs = op.SubOperations
        If Not (subs Is Nothing) Then
            For j = 1 To subs.count
                Dim subop As SubOperation
                Set subop = subs(j)
                ' 获取该子操作使用的刀具对象
                Dim t As MillTool
                Set t = subop.Tool
                If Not (t Is Nothing) Then
                    ' ===== 从子操作名称中提取"加工方式" =====
                    ' 子操作的 Name 通常格式为：
                    '   "精加工   刀具 2   T2 - 2MM Z-6 刀"
                    '   我们需要截取第一个空格前的部分（如"精加工"）
                    Dim procName As String
                    procName = subop.Name
                    Dim spacePos As Integer
                    
                    ' 优先找双空格（"  "）作为分割点
                    spacePos = InStr(procName, "  ")
                    If spacePos > 0 Then
                        ' 取双空格之前的部分作为加工方式
                        procName = Left(procName, spacePos - 1)
                    Else
                        ' 如果没有双空格，退一步找单空格
                        spacePos = InStr(procName, " ")
                        If spacePos > 0 Then
                            procName = Left(procName, spacePos - 1)
                        End If
                    End If
                    
                    ' ===== 组合下拉框显示文本 =====
                    ' 格式："加工方式 - 刀具名称"，例如 "精加工 - Flat10mm"
                    Dim displayText As String
                    displayText = procName & " - " & t.Name
                    
                    ' ===== 去重：如果下拉框中还没有该组合，则添加 =====
                    ' 防止同一把刀具在多个操作中出现导致下拉列表重复
                    Dim found As Boolean: found = False
                    Dim k As Long
                    For k = 0 To cmbTool.ListCount - 1
                        If cmbTool.List(k) = displayText Then
                            found = True
                            Exit For
                        End If
                    Next k
                    If Not found Then cmbTool.AddItem displayText
                End If
            Next j
        End If
    Next i
    
    ' --- 如果有刀具，默认选中第一项 ---
    If cmbTool.ListCount > 0 Then cmbTool.ListIndex = 0
    
    ' --- 设置偏移量的默认值为 0 ---
    txtX.Text = "0"
    txtY.Text = "0"
    txtZ.Text = "0"
End Sub


' ==============================================================================
' cmdOK_Click — 用户点击"确定"按钮
' 功能：验证输入，将选择结果编码到窗体的 Tag 属性中，然后隐藏窗体
' Tag 格式约定： "刀具名称|X偏移|Y偏移|Z偏移"
' 主模块（全排版刀具偏移）读取 Tag 来获取用户的选择
' ==============================================================================
Private Sub cmdOK_Click()
    ' --- 校验：必须选择一把刀具 ---
    If cmbTool.ListIndex < 0 Then
        MsgBox "请选择一把刀具！", vbExclamation
        Exit Sub
    End If
    
    ' --- 读取三个文本框中的数值 ---
    ' Val() 函数会把非数字开头的字符串转为 0，所以不会崩溃
    Dim xVal As Double, yVal As Double, zVal As Double
    xVal = Val(txtX.Text)
    yVal = Val(txtY.Text)
    zVal = Val(txtZ.Text)
    
    ' --- 校验：偏移量不能全为 0（全为 0 则移动没有意义） ---
    If xVal = 0 And yVal = 0 And zVal = 0 Then
        MsgBox "偏移量不能全为 0！", vbExclamation
        Exit Sub
    End If
    
    ' ===== 从下拉框的显示文本中提取出纯刀具名称 =====
    ' 下拉框显示的是 "加工方式 - 刀具名称" 格式（如 "精加工 - Flat10mm"）
    ' 但主模块需要纯刀具名（如 "Flat10mm"）来匹配刀具
    Dim toolName As String
    Dim dashPos As Integer
    ' 查找 " - " 分隔符的位置
    dashPos = InStr(cmbTool.Text, " - ")
    If dashPos > 0 Then
        ' " - " 长度为 3，所以从 dashPos + 3 开始取就是刀具名
        toolName = Mid(cmbTool.Text, dashPos + 3)
    Else
        ' 如果没有" - "（理论上不会），就使用全文
        toolName = cmbTool.Text
    End If
    
    ' --- 将结果编码到窗体的 Tag 属性中 ---
    ' Tag 是一个字符串属性，这里用它来跨窗体传递数据
    ' 格式：刀具名|X偏移|Y偏移|Z偏移
    Me.Tag = toolName & "|" & CStr(xVal) & "|" & CStr(yVal) & "|" & CStr(zVal)
    
    ' --- 隐藏窗体（结束 Show 模态循环） ---
    ' 注意是 Hide 不是 Unload，这样主模块还可以读取 Me.Tag
    Me.Hide
End Sub


' ==============================================================================
' cmdCancel_Click — 用户点击"取消"按钮
' 功能：将 Tag 置为空字符串（主模块据此判断用户取消），然后隐藏窗体
' ==============================================================================
Private Sub cmdCancel_Click()
    ' Tag 置空 → 主模块的 If frmToolOffset.Tag = "" Then Exit Sub 会触发
    Me.Tag = ""
    Me.Hide
End Sub
' ==================== UserForm 代码结束 ====================
