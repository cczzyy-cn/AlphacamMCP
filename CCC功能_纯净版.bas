


Public Sub ApplySortToDrawing(ByRef sortedKeys() As String)
    On Error GoTo ErrHandler3
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    Dim ops As Operations: Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub
    
    Dim opIdx As Long, s As Long, si As Long, sj As Long, mi As Long
    Dim mSheet As Long, pos As Long
    
    ' ===== 先读当前 OpNo（不清空），解码排版 =====
    Dim opNos() As Long: ReDim opNos(1 To ops.count)
    Dim maxOpNo As Long: maxOpNo = 0
    For opIdx = 1 To ops.count
        Dim opR As Operation: Set opR = ops(opIdx)
        Dim sbR As SubOperations: Set sbR = opR.SubOperations
        opNos(opIdx) = 0
        If Not (sbR Is Nothing) Then
            If sbR.count > 0 Then
                Dim subR As SubOperation: Set subR = sbR(1)
                Dim tpR As paths: Set tpR = subR.ToolPaths
                If Not (tpR Is Nothing) Then
                    If tpR.count > 0 Then
                        Dim tR As Path: Set tR = tpR(1)
                        If Not (tR Is Nothing) Then
                            opNos(opIdx) = tR.OpNo
                            If tR.OpNo > maxOpNo Then maxOpNo = tR.OpNo
                        End If
                    End If
                End If
            End If
        End If
    Next opIdx
    
    ' ===== NestInformation =====
    Dim ni As NestInformation: Set ni = drw.GetNestInformation()
    Dim sheetCount As Long: sheetCount = ni.Sheets.count
    If sheetCount = 0 Then sheetCount = 1
    
    ' 排版路径名字典
    Dim shNm() As Object: ReDim shNm(1 To sheetCount)
    For s = 1 To sheetCount
        Dim nd As Object: Set nd = CreateObject("Scripting.Dictionary")
        Dim sp As paths: Set sp = ni.Sheets(s).paths
        For mi = 1 To sp.count
            If Not nd.Exists(sp(mi).Name) Then nd.Add sp(mi).Name, True
        Next mi
        Set shNm(s) = nd
    Next s
    
    ' ===== 匹配 OP 到排版 =====
    Dim shMap As Object: Set shMap = CreateObject("Scripting.Dictionary")
    Dim lastSh As Long: lastSh = 0
    
    For opIdx = 1 To ops.count
        mSheet = 0
        
        ' 1) 如果有 OpNo >= 100，直接解码
        If opNos(opIdx) >= 100 Then
            mSheet = opNos(opIdx) \ 100
            If mSheet > sheetCount Then mSheet = sheetCount
            If mSheet < 1 Then mSheet = 1
        End If
        
        ' 2) 否则用 NestInformation
        If mSheet = 0 Then
            Dim opN As Operation: Set opN = ops(opIdx)
            Dim sbN As SubOperations: Set sbN = opN.SubOperations
            If Not (sbN Is Nothing) Then
                Dim siN As Long
                For siN = 1 To sbN.count
                    Dim subN As SubOperation: Set subN = sbN(siN)
                    Dim tpN As paths: Set tpN = subN.ToolPaths
                    If Not (tpN Is Nothing) Then
                        For mi = 1 To tpN.count
                            Dim tN2 As Path: Set tN2 = tpN(mi)
                            If Not (tN2 Is Nothing) Then
                                For s = 1 To sheetCount
                                    If shNm(s).Exists(tN2.Name) Then
                                        mSheet = s: lastSh = s: Exit For
                                    End If
                                Next s
                                If mSheet > 0 Then Exit For
                            End If
                        Next mi
                    End If
                    If mSheet > 0 Then Exit For
                Next siN
            End If
        End If
        
        If mSheet = 0 And lastSh > 0 Then mSheet = lastSh
        If mSheet = 0 Then mSheet = 1
        shMap.Add opIdx, mSheet
    Next opIdx
    
    ' ===== 按 (排版, 刀具) 分组 =====
    Dim stD As Object: Set stD = CreateObject("Scripting.Dictionary")
    
    For opIdx = 1 To ops.count
        Dim opM As Operation: Set opM = ops(opIdx)
        Dim sbM As SubOperations: Set sbM = opM.SubOperations
        If Not (sbM Is Nothing) Then
            For si = 1 To sbM.count
                Dim subM As SubOperation: Set subM = sbM(si)
                Dim tM As MillTool: Set tM = subM.Tool
                If Not (tM Is Nothing) Then
                    Dim mn As String: mn = subM.Name
                    Dim sp As Integer
                    sp = InStr(mn, "  ")
                    If sp > 0 Then mn = Left(mn, sp - 1) Else: sp = InStr(mn, " "): If sp > 0 Then mn = Left(mn, sp - 1)
                    Dim tk As String: tk = mn & " T" & CStr(tM.Number) & " " & tM.Name
                    Dim sid As Long: If shMap.Exists(opIdx) Then sid = shMap(opIdx) Else sid = 1
                    Dim ck As String: ck = CStr(sid) & "|" & tk
                    
                    Dim tpsM As paths: Set tpsM = subM.ToolPaths
                    If Not (tpsM Is Nothing) Then
                        For mi = 1 To tpsM.count
                            Dim tpM As Path: Set tpM = tpsM(mi)
                            If Not (tpM Is Nothing) Then
                                If Not stD.Exists(ck) Then
                                    Dim nc As Collection: Set nc = New Collection
                                    stD.Add ck, nc
                                End If
                                Dim c2 As Collection: Set c2 = stD(ck)
                                c2.Add tpM
                            End If
                        Next mi
                    End If
                End If
            Next si
        End If
    Next opIdx
    
    If stD.count = 0 Then Exit Sub
    
    ' ===== 分配 OpNo（不清空，直接覆盖）=====
    App.SetUndoCommandName "排版刀具排序"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    
    For si = 1 To sheetCount
        pos = 1
        For sj = 0 To UBound(sortedKeys)
            Dim tky As String: tky = sortedKeys(sj)
            Dim ck2 As String: ck2 = CStr(si) & "|" & tky
            If stD.Exists(ck2) Then
                Dim tc As Collection: Set tc = stD(ck2)
                For mi = 1 To tc.count
                    Dim ta As Path: Set ta = tc(mi)
                    If Not (ta Is Nothing) Then ta.OpNo = si * 100 + pos
                Next mi
                pos = pos + 1
            End If
        Next sj
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
