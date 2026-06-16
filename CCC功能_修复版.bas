' ==============================================================================
' CCC功能 合集 — UserForm 版（修复二次排序失效）
' ==============================================================================
' 修复内容：
'   1. 旧 OpNo 清零（消除残留干扰）
'   2. ScanOperations 方法名截取逻辑修复
'   3. ApplySortToDrawing 改用 tpDict 按 toolKey 分组，不依赖 opIdx/subIdx 索引
'   4. cmdOK_Click 改为保持窗体打开（RefreshOpList），便于连续调整
' ==============================================================================

' ==============================================================================
' 功能1：依边界裁剪
' ==============================================================================
Sub 依边界裁剪()
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！", vbExclamation, "依边界裁剪": Exit Sub
    App.SetUndoCommandName "依边界裁剪"
    App.SetUndoPoint
    Dim boundary As Path: Set boundary = drw.UserSelectOneGeo("【依边界裁剪】请选择边界几何图形（封闭图形）")
    If boundary Is Nothing Then Exit Sub
    If Not boundary.Closed Then If MsgBox("边界不是封闭图形，继续吗？", vbYesNo + vbQuestion, "依边界裁剪") = vbNo Then Exit Sub
    drw.SetGeosSelected False
    If Not drw.UserSelectMultiGeos("【依边界裁剪】请选择要裁剪的线段（框选或点选）", 0) Then Exit Sub
    On Error GoTo ErrHandler1
    drw.ScreenUpdating = False
    Dim selectedGeos As New Collection
    Dim g As Path
    For Each g In drw.Geometries
        If g.Selected And (Not (g Is boundary)) Then selectedGeos.Add g
    Next g
    drw.SetGeosSelected False
    boundary.Selected = True
    Dim totalTrimmed As Long
    Dim i As Long
    For i = 1 To selectedGeos.count
        Dim geo As Path: Set geo = selectedGeos(i)
        If geo.Closed Then GoTo SkipGeo1
        If geo.TestIntersectPath(boundary, 0, 0) Then
            Dim firstElem As Element: Set firstElem = geo.GetFirstElem
            Dim lastElem As Element: Set lastElem = geo.GetLastElem
            If firstElem Is Nothing Or lastElem Is Nothing Then GoTo SkipGeo1
            Dim sx As Double, sy As Double, ex As Double, ey As Double
            sx = firstElem.StartXG: sy = firstElem.StartYG
            ex = lastElem.EndXG: ey = lastElem.EndYG
            If Not boundary.IsPointInside(sx, sy) Then geo.TrimWithCuttingGeos sx, sy: totalTrimmed = totalTrimmed + 1
            If Not boundary.IsPointInside(ex, ey) Then geo.TrimWithCuttingGeos ex, ey: totalTrimmed = totalTrimmed + 1
        End If
SkipGeo1:
    Next i
    drw.SetGeosSelected False
    drw.ScreenUpdating = True
    drw.Redraw: drw.ZoomAll
    MsgBox "裁剪完成！已处理 " & totalTrimmed & " 次裁剪。", vbInformation, "依边界裁剪"
    Exit Sub
ErrHandler1:
    drw.ScreenUpdating = True: drw.SetGeosSelected False: drw.Redraw
    MsgBox "裁剪过程中发生错误：" & Err.Description, vbCritical, "依边界裁剪"
End Sub

' ==============================================================================
' 功能2：全排版刀具偏移
' ==============================================================================
Sub 全排版刀具偏移()
    frmToolOffset.Show vbModeless
End Sub

Public Sub ApplyToolOffset(ByVal selectedTool As String, ByVal xOff As Double, ByVal yOff As Double, ByVal zOff As Double)
    On Error GoTo ErrHandler2
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    App.SetUndoCommandName "全排版刀具偏移"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    Dim count As Long
    Dim ops As Operations: Set ops = drw.Operations
    Dim i As Long, j As Long
    For i = 1 To ops.count
        Dim op As Operation: Set op = ops(i)
        Dim subs As SubOperations: Set subs = op.SubOperations
        If Not (subs Is Nothing) Then
            For j = 1 To subs.count
                Dim subop As SubOperation: Set subop = subs(j)
                Dim t As MillTool: Set t = subop.Tool
                If Not (t Is Nothing) Then
                    If t.Name = selectedTool Then
                        Dim tps As Paths: Set tps = subop.ToolPaths
                        If Not (tps Is Nothing) Then
                            Dim m As Long
                            For m = 1 To tps.count
                                Dim tp As Path: Set tp = tps(m)
                                If Not (tp Is Nothing) Then tp.MoveG xOff, yOff, zOff: count = count + 1
                            Next m
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    drw.ScreenUpdating = True: drw.Redraw
    MsgBox "偏移完成！已处理 " & count & " 条刀具路径。" & vbCrLf & "刀具: " & selectedTool & vbCrLf & "偏移量: X=" & xOff & "  Y=" & yOff & "  Z=" & zOff, vbInformation, "全排版刀具偏移"
    Exit Sub
ErrHandler2:
    drw.ScreenUpdating = True
    MsgBox "偏移出错：" & Err.Description, vbCritical
End Sub

' ==============================================================================
' 功能3：排版刀具排序
' ==============================================================================
Sub 排版刀具排序()
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！", vbExclamation, "排版刀具排序": Exit Sub
    If drw.Operations Is Nothing Or drw.Operations.count = 0 Then MsgBox "图纸中没有找到任何操作！", vbExclamation, "排版刀具排序": Exit Sub
    frmToolSort.Show vbModeless
End Sub

' ==============================================================================
' 公有辅助函数
' ==============================================================================

' ★ 修复：ScanOperations — 方法名截取逻辑修复
Public Function ScanOperations() As Collection
    Dim result As New Collection
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Set ScanOperations = result: Exit Function
    Dim ops As Operations: Set ops = drw.Operations
    If ops Is Nothing Then Set ScanOperations = result: Exit Function
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, j As Long
    For i = 1 To ops.count
        Dim op As Operation: Set op = ops(i)
        Dim subs As SubOperations: Set subs = op.SubOperations
        If subs Is Nothing Then GoTo NextOp2
        For j = 1 To subs.count
            Dim subop As SubOperation: Set subop = subs(j)
            Dim t As MillTool: Set t = subop.Tool
            If Not (t Is Nothing) Then
                Dim methodName As String: methodName = subop.Name
                ' ★ 修复：先查双空格，没双空格再查单空格，不串行执行
                Dim spPos As Integer
                spPos = InStr(methodName, "  ")
                If spPos = 0 Then spPos = InStr(methodName, " ")
                If spPos > 0 Then methodName = Left(methodName, spPos - 1)
                
                Dim key As String: key = methodName & " T" & CStr(t.Number) & " " & t.Name
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        Next j
NextOp2:
    Next i
    Dim keysArr: keysArr = dict.Keys
    Dim kk As Long
    For kk = 0 To UBound(keysArr)
        result.Add keysArr(kk)
    Next kk
    Set ScanOperations = result
End Function

' ★ 修复：ApplySortToDrawing — 重写，不再依赖 opIdx/subIdx，改用 tpDict
Public Sub ApplySortToDrawing(ByRef sortedKeys() As String)
    On Error GoTo ErrHandler3
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    
    ' 重新获取 ops（确保最新）
    Dim ops As Operations: Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub
    
    ' ★ 修复1：先清空所有 ToolPath 的 OpNo（消除旧值残留）
    Dim opX As Long, subX As Long, tpX As Long
    For opX = 1 To ops.count
        Dim opClear As Operation: Set opClear = ops(opX)
        Dim subsClear As SubOperations: Set subsClear = opClear.SubOperations
        If Not (subsClear Is Nothing) Then
            For subX = 1 To subsClear.count
                Dim subClear As SubOperation: Set subClear = subsClear(subX)
                Dim tpsClear As Paths: Set tpsClear = subClear.ToolPaths
                If Not (tpsClear Is Nothing) Then
                    For tpX = 1 To tpsClear.count
                        Dim tpClear As Path: Set tpClear = tpsClear(tpX)
                        If Not (tpClear Is Nothing) Then tpClear.OpNo = 0
                    Next tpX
                End If
            Next subX
        End If
    Next opX
    
    ' ★ 修复2：清零后重新获取 ops（确保引用有效）
    Set ops = drw.Operations
    
    ' ★ 修复3：用字典按 toolKey 分组收集工具路径引用（不依赖 opIdx/subIdx）
    Dim tpDict As Object: Set tpDict = CreateObject("Scripting.Dictionary")
    Dim opIdx As Long, subIdx As Long
    
    For opIdx = 1 To ops.count
        Dim opN As Operation: Set opN = ops(opIdx)
        Dim subsN As SubOperations: Set subsN = opN.SubOperations
        If Not (subsN Is Nothing) Then
            For subIdx = 1 To subsN.count
                Dim subN As SubOperation: Set subN = subsN(subIdx)
                Dim tN As MillTool: Set tN = subN.Tool
                If Not (tN Is Nothing) Then
                    Dim mName As String: mName = subN.Name
                    Dim sp As Integer
                    sp = InStr(mName, "  ")
                    If sp = 0 Then sp = InStr(mName, " ")
                    If sp > 0 Then mName = Left(mName, sp - 1)
                    
                    Dim toolKey As String: toolKey = mName & " T" & CStr(tN.Number) & " " & tN.Name
                    
                    ' 收集此子操作的所有工具路径
                    Dim tps As Paths: Set tps = subN.ToolPaths
                    If Not (tps Is Nothing) Then
                        Dim mm As Long
                        For mm = 1 To tps.count
                            Dim tp As Path: Set tp = tps(mm)
                            If Not (tp Is Nothing) Then
                                If Not tpDict.Exists(toolKey) Then
                                    tpDict.Add toolKey, New Collection
                                End If
                                tpDict(toolKey).Add tp
                            End If
                        Next mm
                    End If
                End If
            Next subIdx
        End If
    Next opIdx
    
    If tpDict.count = 0 Then Exit Sub
    
    ' 检测排版数量
    Dim sheetGeos As New Collection
    Dim gi As Long
    For gi = 1 To drw.Geometries.count
        If drw.Geometries(gi).Closed And drw.Geometries(gi).Sheet Then
            sheetGeos.Add drw.Geometries(gi)
        End If
    Next gi
    Dim sheetCount As Long: sheetCount = sheetGeos.count
    If sheetCount = 0 Then sheetCount = 1
    
    App.SetUndoCommandName "排版刀具排序"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    
    ' ★ 修复4：按排版逐张分配 OpNo
    Dim baseOpNo As Long: baseOpNo = 1
    Dim sheetIdx As Long
    
    For sheetIdx = 1 To sheetCount
        Dim pos As Long: pos = 0
        Dim sj As Long
        
        ' 按用户排序的顺序分配
        For sj = 0 To UBound(sortedKeys)
            Dim targetKey As String: targetKey = sortedKeys(sj)
            
            If tpDict.Exists(targetKey) Then
                Dim tpCol As Collection: Set tpCol = tpDict(targetKey)
                
                ' 此刀具类型的路径按 Sheet 数均分
                Dim perSheet As Long: perSheet = tpCol.count \ sheetCount
                Dim extra As Long: extra = tpCol.count Mod sheetCount
                
                ' 计算当前 Sheet 在此刀具中的起止索引
                Dim startIdx As Long: startIdx = 1
                Dim s As Long
                For s = 1 To sheetIdx - 1
                    Dim prevSize As Long: prevSize = perSheet
                    If s <= extra Then prevSize = prevSize + 1
                    startIdx = startIdx + prevSize
                Next s
                
                Dim sizeThis As Long: sizeThis = perSheet
                If sheetIdx <= extra Then sizeThis = sizeThis + 1
                Dim endIdx As Long: endIdx = startIdx + sizeThis - 1
                If endIdx > tpCol.count Then endIdx = tpCol.count
                
                ' 给当前排版中此刀具的路径分配连续 OpNo
                Dim t As Long
                For t = startIdx To endIdx
                    Dim tpAssign As Path: Set tpAssign = tpCol(t)
                    If Not (tpAssign Is Nothing) Then
                        tpAssign.OpNo = baseOpNo + pos
                        pos = pos + 1
                    End If
                Next t
            End If
        Next sj
        
        baseOpNo = baseOpNo + pos
    Next sheetIdx
    
    ' ★ 修复5：重新获取 ops 再调用 OrderAll（确保引用有效）
    Set ops = drw.Operations
    ops.OrderAll
    
    drw.ScreenUpdating = True: drw.Redraw
    Exit Sub
    
ErrHandler3:
    drw.ScreenUpdating = True
    MsgBox "排版刀具排序出错：" & Err.Description, vbCritical
End Sub

' ==============================================================================
' frmToolSort 窗体代码（修复 cmdOK_Click 保持窗体打开）
' ==============================================================================
' ★ 注意：以下为窗体代码，需要放在 frmToolSort 的代码窗口中
' ★ 在 VBA 编辑器中双击 "frmToolSort" 窗体，粘贴以下代码

'Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Const GWL_STYLE As Long = (-16)
'Private Const WS_MINIMIZEBOX As Long = &H20000
'Private Const SWP_NOSIZE As Long = &H1
'Private Const SWP_NOMOVE As Long = &H2
'Private Const SWP_NOZORDER As Long = &H4
'Private Const SWP_FRAMECHANGED As Long = &H20
'
'Private Sub AddMinButton(frm As Object)
'    Dim hWnd As Long
'    hWnd = FindWindow("ThunderDFrame", frm.Caption)
'    If hWnd = 0 Then hWnd = FindWindow("ThunderRT5Form", frm.Caption)
'    If hWnd = 0 Then Exit Sub
'    Dim style As Long
'    style = GetWindowLong(hWnd, GWL_STYLE)
'    style = style Or WS_MINIMIZEBOX
'    SetWindowLong hWnd, GWL_STYLE, style
'    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
'End Sub
'
'Private Sub UserForm_Activate(): AddMinButton Me: End Sub
'
'Private Sub UserForm_Initialize()
'    Me.Caption = "排版刀具排序"
'    RefreshOpList
'End Sub
'
'Private Sub cmdRefresh_Click(): RefreshOpList: End Sub
'
'Public Sub RefreshOpList()
'    lstTools.Clear
'    Dim rawList As Collection: Set rawList = ScanOperations()
'    If rawList Is Nothing Or rawList.count = 0 Then lstTools.AddItem "(未找到操作)": Exit Sub
'    Dim k As Long
'    For k = 1 To rawList.count: lstTools.AddItem rawList(k): Next k
'    If lstTools.ListCount > 0 Then lstTools.ListIndex = 0
'End Sub
'
'Private Sub cmdUp_Click()
'    If lstTools.ListIndex < 1 Then Exit Sub
'    Dim idx As Long: idx = lstTools.ListIndex
'    Dim txt As String: txt = lstTools.List(idx)
'    lstTools.RemoveItem idx: lstTools.AddItem txt, idx - 1
'    lstTools.ListIndex = idx - 1
'End Sub
'
'Private Sub cmdDown_Click()
'    If lstTools.ListIndex < 0 Then Exit Sub
'    If lstTools.ListIndex >= lstTools.ListCount - 1 Then Exit Sub
'    Dim idx As Long: idx = lstTools.ListIndex
'    Dim txt As String: txt = lstTools.List(idx)
'    lstTools.RemoveItem idx: lstTools.AddItem txt, idx + 1
'    lstTools.ListIndex = idx + 1
'End Sub
'
' ★ 修复：cmdOK_Click 不再 Unload Me，改为刷新列表，保持窗体打开
'Private Sub cmdOK_Click()
'    If lstTools.ListCount = 0 Or lstTools.List(0) = "(未找到操作)" Then
'        MsgBox "没有可排序的操作！请先点击【刷新】扫描图纸。", vbExclamation
'        Exit Sub
'    End If
'    Dim count As Long: count = lstTools.ListCount
'    ReDim sortedKeys(0 To count - 1) As String
'    Dim i As Long
'    For i = 0 To count - 1: sortedKeys(i) = lstTools.List(i): Next i
'    ApplySortToDrawing sortedKeys
'    MsgBox "排序已应用！共 " & count & " 个加工步骤。", vbInformation, "排版刀具排序"
'    ' ★ 不关闭窗体，刷新列表便于再次调整
'    RefreshOpList
'End Sub
'
'Private Sub cmdClose_Click(): Unload Me: End Sub


' ==============================================================================
' 菜单注册（不变）
' ==============================================================================
Function InitAlphacamAddIn(AcamVersion As Long) As Integer
    Dim frm As Frame: Set frm = App.Frame
    With frm
        Dim barId As Long: barId = .CreateButtonBar("CCC功能")
        .AddMenuItem3 "依边界裁&剪", "m_依边界裁剪", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "cut.bmp", .LastMenuCommandID
        .AddMenuItem3 "全排版刀具偏&移", "m_全排版刀具偏移", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "offset.bmp", .LastMenuCommandID
        .AddMenuItem3 "排版刀具排&序", "m_排版刀具排序", acamMenuNEW, "CCC功能", vbNullString
        .AddButton barId, "sort.bmp", .LastMenuCommandID
    End With
    InitAlphacamAddIn = 0
End Function

Function m_依边界裁剪(): 依边界裁剪: End Function
Function m_全排版刀具偏移(): 全排版刀具偏移: End Function
Function m_排版刀具排序(): 排版刀具排序: End Function
