Dim g_sheetMap As Object

Public Sub ApplySortToDrawing(ByRef sortedKeys() As String)
    On Error GoTo ErrHandler3
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    Dim ops As Operations: Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub
    
    Dim opIdx As Long
    
    ' ===== 首次排序：用 NestInformation 匹配并保存映射 =====
    If g_sheetMap Is Nothing Then
        Dim ni As NestInformation: Set ni = drw.GetNestInformation()
        Dim sheetCount As Long: sheetCount = ni.Sheets.count
        If sheetCount = 0 Then sheetCount = 1
        
        ' 排版路径名字典
        Dim sheetNames() As Object: ReDim sheetNames(1 To sheetCount)
        Dim s As Long
        For s = 1 To sheetCount
            Dim nd As Object: Set nd = CreateObject("Scripting.Dictionary")
            Dim sps As paths: Set sps = ni.Sheets(s).paths
            Dim pp As Long
            For pp = 1 To sps.count
                If Not nd.Exists(sps(pp).Name) Then nd.Add sps(pp).Name, True
            Next pp
            Set sheetNames(s) = nd
        Next s
        
        Set g_sheetMap = CreateObject("Scripting.Dictionary")
        
        For opIdx = 1 To ops.count
            Dim opN As Operation: Set opN = ops(opIdx)
            Dim subsN As SubOperations: Set subsN = opN.SubOperations
            Dim mSheet As Long: mSheet = 0
            
            If Not (subsN Is Nothing) Then
                Dim subN1 As SubOperation
                For Each subN1 In subsN
                    Dim tpsN1 As paths: Set tpsN1 = subN1.ToolPaths
                    If Not (tpsN1 Is Nothing) Then
                        Dim tpN1 As Path
                        For Each tpN1 In tpsN1
                            For s = 1 To sheetCount
                                If sheetNames(s).Exists(tpN1.Name) Then
                                    mSheet = s
                                    Exit For
                                End If
                            Next s
                            If mSheet > 0 Then Exit For
                        Next tpN1
                    End If
                    If mSheet > 0 Then Exit For
                Next subN1
            End If
            
            If mSheet = 0 Then
                Dim pv As Long
                For pv = opIdx - 1 To 1 Step -1
                    If g_sheetMap.Exists(pv) Then
                        mSheet = g_sheetMap(pv)
                        Exit For
                    End If
                Next pv
            End If
            If mSheet = 0 Then mSheet = 1
            
            g_sheetMap.Add opIdx, mSheet
        Next opIdx
        
        sheetCount = 0
        Dim uniqS As Object: Set uniqS = CreateObject("Scripting.Dictionary")
        For opIdx = 1 To ops.count
            If g_sheetMap.Exists(opIdx) Then
                If Not uniqS.Exists(g_sheetMap(opIdx)) Then
                    uniqS.Add g_sheetMap(opIdx), True
                    sheetCount = sheetCount + 1
                End If
            End If
        Next opIdx
        If sheetCount = 0 Then sheetCount = 1
    Else
        ' 后续排序：从已有映射获取排版数
        Dim uniq As Object: Set uniq = CreateObject("Scripting.Dictionary")
        For opIdx = 1 To ops.count
            If g_sheetMap.Exists(opIdx) Then
                If Not uniq.Exists(g_sheetMap(opIdx)) Then
                    uniq.Add g_sheetMap(opIdx), True
                End If
            End If
        Next opIdx
        sheetCount = uniq.count
        If sheetCount = 0 Then sheetCount = 1
    End If
    
    ' ===== 按 (排版, 刀具) 分组 =====
    Dim stDict As Object: Set stDict = CreateObject("Scripting.Dictionary")
    
    For opIdx = 1 To ops.count
        Dim opM As Operation: Set opM = ops(opIdx)
        Dim subsM As SubOperations: Set subsM = opM.SubOperations
        If Not (subsM Is Nothing) Then
            Dim si As Long
            For si = 1 To subsM.count
                Dim subM As SubOperation: Set subM = subsM(si)
                Dim tM As MillTool: Set tM = subM.Tool
                If Not (tM Is Nothing) Then
                    Dim mn As String: mn = subM.Name
                    Dim sp As Integer
                    sp = InStr(mn, "  ")
                    If sp > 0 Then mn = Left(mn, sp - 1) Else: sp = InStr(mn, " "): If sp > 0 Then mn = Left(mn, sp - 1)
                    Dim tk As String: tk = mn & " T" & CStr(tM.Number) & " " & tM.Name
                    Dim sid As Long: If g_sheetMap.Exists(opIdx) Then sid = g_sheetMap(opIdx) Else sid = 1
                    Dim ck As String: ck = CStr(sid) & "|" & tk
                    
                    Dim tpsM As paths: Set tpsM = subM.ToolPaths
                    If Not (tpsM Is Nothing) Then
                        Dim mi As Long
                        For mi = 1 To tpsM.count
                            Dim tpM As Path: Set tpM = tpsM(mi)
                            If Not (tpM Is Nothing) Then
                                If Not stDict.Exists(ck) Then
                                    Dim nc As Collection: Set nc = New Collection
                                    stDict.Add ck, nc
                                End If
                                Dim c2 As Collection: Set c2 = stDict(ck)
                                c2.Add tpM
                            End If
                        Next mi
                    End If
                End If
            Next si
        End If
    Next opIdx
    
    If stDict.count = 0 Then Exit Sub
    
    ' ===== 清空旧 OpNo =====
    Dim ox As Long, sx As Long, tx As Long
    For ox = 1 To ops.count
        Dim oc As Operation: Set oc = ops(ox)
        Dim sc As SubOperations: Set sc = oc.SubOperations
        If Not (sc Is Nothing) Then
            For sx = 1 To sc.count
                Dim sbc As SubOperation: Set sbc = sc(sx)
                Dim tpc As paths: Set tpc = sbc.ToolPaths
                If Not (tpc Is Nothing) Then
                    For tx = 1 To tpc.count
                        Dim tpc2 As Path: Set tpc2 = tpc(tx)
                        If Not (tpc2 Is Nothing) Then tpc2.OpNo = 0
                    Next tx
                End If
            Next sx
        End If
    Next ox
    
    ' ===== 分配 OpNo =====
    App.SetUndoCommandName "排版刀具排序"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    
    Dim b As Long: b = 1
    For si = 1 To sheetCount
        Dim pos As Long: pos = 0
        Dim sj As Long
        For sj = 0 To UBound(sortedKeys)
            Dim tky As String: tky = sortedKeys(sj)
            Dim ck2 As String: ck2 = CStr(si) & "|" & tky
            If stDict.Exists(ck2) Then
                Dim tc As Collection: Set tc = stDict(ck2)
                Dim tt As Long
                For tt = 1 To tc.count
                    Dim ta As Path: Set ta = tc(tt)
                    If Not (ta Is Nothing) Then ta.OpNo = b + pos
                Next tt
                pos = pos + 1
            End If
        Next sj
        b = b + pos
    Next si
    
    ops.OrderAll
    drw.ScreenUpdating = True: drw.Redraw
    Exit Sub
    
ErrHandler3:
    drw.ScreenUpdating = True
    MsgBox "排版刀具排序出错：" & Err.Description, vbCritical
End Sub

' ==============================================================================
' CCC功能 合集 — 最终版
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
                        Dim tps As paths: Set tps = subop.ToolPaths
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
' 公共辅助函数
' ==============================================================================
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
                Dim spPos As Integer
                spPos = InStr(methodName, "  ")
                If spPos > 0 Then methodName = Left(methodName, spPos - 1) Else: spPos = InStr(methodName, " "): If spPos > 0 Then methodName = Left(methodName, spPos - 1)
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

' ★ 最终版：原始算法 + OpNo清零（仅修复二次排序失效）


' ★ 全局排序版：按刀具顺序全局分配OpNo，不按排版分组













' ==============================================================================
' 菜单注册
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
