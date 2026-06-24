' ==============================================================================
' 全排版刀具偏移 — 对排版内所有使用某把刀具的路径进行 XYZ 偏移
'
' 功能：
'   1. 自动扫描当前图纸所有操作，列出使用的刀具
'   2. 用户选择要偏移的刀具
'   3. 输入 X、Y、Z 偏移量
'   4. 自动对所有使用该刀具的路径执行 MoveG(x, y, z)
'
' 安装：
'   在 VBA 编辑器中插入模块，粘贴此代码
'   保存为 .arb 放入 LICOMDIR\VBMacros\StartUp\ 自动注册菜单
' ==============================================================================

' 手动运行时从此入口
Sub 全排版刀具偏移()
    Dim drw As Drawing
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        MsgBox "没有活动图纸！", vbExclamation, "全排版刀具偏移"
        Exit Sub
    End If
    
    App.SetUndoCommandName "全排版刀具偏移"
    App.SetUndoPoint
    
    ' === 步骤 1：收集图纸中所有使用到的刀具 ===
    Dim toolList As New Collection
    Dim ops As Operations
    Set ops = drw.Operations
    
    If ops Is Nothing Or ops.Count = 0 Then
        MsgBox "图纸中没有找到任何操作！", vbExclamation, "全排版刀具偏移"
        Exit Sub
    End If
    
    Dim i As Long, j As Long
    For i = 1 To ops.Count
        Dim op As Operation
        Set op = ops(i)
        Dim subs As SubOperations
        Set subs = op.SubOperations
        If Not (subs Is Nothing) Then
            For j = 1 To subs.Count
                Dim subop As SubOperation
                Set subop = subs(j)
                Dim t As MillTool
                Set t = subop.Tool
                If Not (t Is Nothing) Then
                    ' 去重
                    Dim found As Boolean: found = False
                    Dim k As Long
                    For k = 1 To toolList.Count
                        If toolList(k) = t.Name Then
                            found = True
                            Exit For
                        End If
                    Next k
                    If Not found Then
                        toolList.Add t.Name
                    End If
                End If
            Next j
        End If
    Next i
    
    If toolList.Count = 0 Then
        MsgBox "没有找到任何刀具！", vbExclamation, "全排版刀具偏移"
        Exit Sub
    End If
    
    ' === 步骤 2：显示刀具列表供选择 ===
    Dim msg As String
    msg = "当前排版使用以下刀具：" & vbCrLf & vbCrLf
    For i = 1 To toolList.Count
        msg = msg & "  " & i & ". " & toolList(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "请输入要偏移的刀具名称（完整名称）："
    
    Dim selectedTool As String
    selectedTool = InputBox(msg, "全排版刀具偏移 - 选择刀具")
    If selectedTool = "" Then Exit Sub
    
    ' === 步骤 3：输入 XYZ 偏移量 ===
    Dim xOff As Double, yOff As Double, zOff As Double
    Dim inputStr As String
    
    inputStr = InputBox("请输入 X 偏移量（正数向右）：", "全排版刀具偏移 - X偏移", "0")
    If inputStr = "" Then Exit Sub
    xOff = Val(inputStr)
    
    inputStr = InputBox("请输入 Y 偏移量（正数向上）：", "全排版刀具偏移 - Y偏移", "0")
    If inputStr = "" Then Exit Sub
    yOff = Val(inputStr)
    
    inputStr = InputBox("请输入 Z 偏移量（正数向上）：", "全排版刀具偏移 - Z偏移", "0")
    If inputStr = "" Then Exit Sub
    zOff = Val(inputStr)
    
    If xOff = 0 And yOff = 0 And zOff = 0 Then
        MsgBox "偏移量不能全为 0！", vbExclamation, "全排版刀具偏移"
        Exit Sub
    End If
    
    ' === 步骤 4：执行偏移 ===
    drw.ScreenUpdating = False
    Dim count As Long: count = 0
    
    For i = 1 To ops.Count
        Set op = ops(i)
        Set subs = op.SubOperations
        If Not (subs Is Nothing) Then
            For j = 1 To subs.Count
                Set subop = subs(j)
                Set t = subop.Tool
                If Not (t Is Nothing) Then
                    If t.Name = selectedTool Then
                        ' 偏移该子操作的所有刀具路径
                        Dim tps As Paths
                        Set tps = subop.ToolPaths
                        If Not (tps Is Nothing) Then
                            Dim m As Long
                            For m = 1 To tps.Count
                                Dim tp As Path
                                Set tp = tps(m)
                                If Not (tp Is Nothing) Then
                                    tp.MoveG xOff, yOff, zOff
                                    count = count + 1
                                End If
                            Next m
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    
    drw.ScreenUpdating = True
    drw.Redraw
    
    MsgBox "偏移完成！已处理 " & count & " 条刀具路径。" & vbCrLf & _
           "偏移量: X=" & xOff & "  Y=" & yOff & "  Z=" & zOff, _
           vbInformation, "全排版刀具偏移"
End Sub

' ==============================================================================
' 自动注册到 CCC功能 菜单
' ==============================================================================
Function InitAlphacamAddIn(AcamVersion As Long) As Integer
    Dim frm As Frame
    Set frm = App.Frame
    With frm
        Dim barId As Long
        barId = .CreateButtonBar("CCC功能")
        
        ' 添加到 CCC功能 顶级菜单下
        .AddMenuItem3 "全排版刀具偏&移", "m_全排版刀具偏移", acamMenuNEW, "CCC功能", vbNullString
        ' 添加到工具栏
        .AddButton barId, "CCC功能.bmp", .LastMenuCommandID
    End With
    InitAlphacamAddIn = 0
End Function

' 回调函数
Function m_全排版刀具偏移()
    全排版刀具偏移
End Function
