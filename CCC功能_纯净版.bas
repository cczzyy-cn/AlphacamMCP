' ==============================================================================

' CCC功能 合集 — UserForm 版（原始算法 + OpNo清零修复）

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

' 公有辅助函数（原始算法 + OpNo清零修复）

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



' ★ 唯一修复：原始算法 + OpNo清零

' ==============================================================================
' ApplySortToDrawing — 重写版，使用 NestInformation 精确匹配排版
' ==============================================================================
' 解决：每个排版的 OP 数量不同，且 OrderAll 会拆分多 SubOp 的 OP
' 方法：用 NestInformation.Sheets 路径名匹配 T2，位置推断 T3
' ==============================================================================
' ==============================================================================
' ApplySortToDrawing — 用 OpNo 值推断排版归属（不依赖 NestInformation）
' ==============================================================================
' 排序前的 OpNo 值（或上一次排序的 OpNo）已经隐含了排版分组信息。
' 同一排版的 OP 具有连续的 OpNo，检测 OpNo 变化趋势即可重建分组。
' ==============================================================================
Public Sub ApplySortToDrawing(ByRef sortedKeys() As String)
    On Error GoTo ErrHandler3
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    
    Dim ops As Operations: Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub
    
    ' ===== 阶段0：从 OpNo 推断排版分组 =====
    ' 先读取每个 OP 第一个 ToolPath 的当前 OpNo（作为该 OP 的 OpNo）
    Dim opNos() As Long: ReDim opNos(1 To ops.count)
    Dim opIdx As Long, subIdx As Long
    For opIdx = 1 To ops.count
        Dim op0 As Operation: Set op0 = ops(opIdx)
        Dim subs0 As SubOperations: Set subs0 = op0.SubOperations
        opNos(opIdx) = 0
        If Not (subs0 Is Nothing) Then
            If subs0.count > 0 Then
                Dim sub0 As SubOperation: Set sub0 = subs0(1)
                Dim tps0 As paths: Set tps0 = sub0.ToolPaths
                If Not (tps0 Is Nothing) Then
                    If tps0.count > 0 Then
                        Dim tp0 As Path: Set tp0 = tps0(1)
                        If Not (tp0 Is Nothing) Then opNos(opIdx) = tp0.OpNo
                    End If
                End If
            End If
        End If
    Next opIdx
    
    ' 如果 OpNo 全为 0（首次排序），用位置推断排版
    Dim allZero As Boolean: allZero = True
    For opIdx = 1 To ops.count
        If opNos(opIdx) <> 0 Then allZero = False: Exit For
    Next opIdx
    
    ' 检测排版数量（从几何或固定值）
    Dim sheetCount As Long: sheetCount = 0
    Dim gi As Long
    For gi = 1 To drw.Geometries.count
        If drw.Geometries(gi).Closed And drw.Geometries(gi).Sheet Then sheetCount = sheetCount + 1
    Next gi
    ' 如果geometry没检测到，尝试从NestInformation获取
    If sheetCount = 0 Then
        On Error Resume Next
        Dim niTmp As Object: Set niTmp = drw.GetNestInformation()
        If Not (niTmp Is Nothing) Then sheetCount = niTmp.Sheets.count
        On Error GoTo ErrHandler3
    End If
    If sheetCount = 0 Then sheetCount = 1
    
    ' 构建 OP→Sheet 映射
    Dim opSheetMap As Object: Set opSheetMap = CreateObject("Scripting.Dictionary")
    
    If allZero Then
        ' 首次排序：按位置均分
        Dim opsPerSheet As Long: opsPerSheet = (ops.count + sheetCount - 1) \ sheetCount
        For opIdx = 1 To ops.count
            Dim inferredSheet As Long: inferredSheet = ((opIdx - 1) \ opsPerSheet) + 1
            If inferredSheet > sheetCount Then inferredSheet = sheetCount
            opSheetMap.Add opIdx, inferredSheet
        Next opIdx
    Else
        ' 非首次排序：用 OpNo 分组，连续 OpNo 属于同一排版
        ' 先按 OpNo 排序 OP
        Dim sortedOps As Object: Set sortedOps = CreateObject("Scripting.Dictionary")
        For opIdx = 1 To ops.count
            If opNos(opIdx) > 0 Then
                If Not sortedOps.Exists(opNos(opIdx)) Then
                    sortedOps.Add opNos(opIdx), opIdx
                End If
            End If
        Next opIdx
        
        Dim sortedOpNos: sortedOpNos = sortedOps.Keys
        ' 对 keys 排序（VBA Dictionary 不保证顺序，需要冒泡排序）
        Dim si As Long, sj As Long
        For si = 0 To UBound(sortedOpNos) - 1
            For sj = si + 1 To UBound(sortedOpNos)
                If sortedOpNos(si) > sortedOpNos(sj) Then
                    Dim tmp As Long: tmp = sortedOpNos(si)
                    sortedOpNos(si) = sortedOpNos(sj)
                    sortedOpNos(sj) = tmp
                End If
            Next sj
        Next si
        
        ' 按 OpNo 顺序分组，检测 OpNo 不连续的地方作为排版边界
        Dim currentSheet As Long: currentSheet = 1
        If UBound(sortedOpNos) >= 0 Then
            Dim prevOpNo As Long: prevOpNo = sortedOpNos(0)
            Dim firstOpIdx As Long: firstOpIdx = sortedOps(sortedOpNos(0))
            opSheetMap.Add firstOpIdx, currentSheet
            
            For si = 1 To UBound(sortedOpNos)
                Dim curOpNo As Long: curOpNo = sortedOpNos(si)
                Dim curOpIdx As Long: curOpIdx = sortedOps(curOpNo)
                
                ' 如果 OpNo 不连续（跳跃 > 1），认为是新排版
                If curOpNo > prevOpNo + 1 Then
                    currentSheet = currentSheet + 1
                    If currentSheet > sheetCount Then currentSheet = sheetCount
                End If
                
                opSheetMap.Add curOpIdx, currentSheet
                prevOpNo = curOpNo
            Next si
        End If
        
        ' 覆盖未匹配的 OP（OpNo=0 的），分到最近有值的排版
        For opIdx = 1 To ops.count
            If Not opSheetMap.Exists(opIdx) Then
                ' 找最近的已匹配 OP
                Dim nearestSheet As Long: nearestSheet = 1
                Dim dist As Long: dist = 9999
                Dim oi As Long
                For oi = 1 To ops.count
                    If opSheetMap.Exists(oi) Then
                        Dim d As Long: d = Abs(oi - opIdx)
                        If d < dist Then dist = d: nearestSheet = opSheetMap(oi)
                    End If
                Next oi
                opSheetMap.Add opIdx, nearestSheet
            End If
        Next opIdx
    End If
    
    ' ===== 阶段1：构建 mergeDict =====
    Dim mergeDict As Object: Set mergeDict = CreateObject("Scripting.Dictionary")
    
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
                    If sp > 0 Then mName = Left(mName, sp - 1) Else: sp = InStr(mName, " "): If sp > 0 Then mName = Left(mName, sp - 1)
                    
                    Dim toolKey As String: toolKey = mName & " T" & CStr(tN.Number) & " " & tN.Name
                    Dim dictKey As String: dictKey = CStr(opIdx) & "|" & toolKey
                    
                    If mergeDict.Exists(dictKey) Then
                        mergeDict(dictKey) = mergeDict(dictKey) & "," & CStr(subIdx)
                    Else
                        mergeDict.Add dictKey, CStr(subIdx)
                    End If
                End If
            Next subIdx
        End If
    Next opIdx
    
    Dim allSubOps As New Collection
    Dim keysArr: keysArr = mergeDict.Keys
    Dim ki As Long
    For ki = 0 To UBound(keysArr)
        allSubOps.Add keysArr(ki) & "|" & mergeDict(keysArr(ki))
    Next ki
    If allSubOps.count = 0 Then Exit Sub
    
    ' ===== 阶段2：清空旧 OpNo =====
    Dim opX As Long, subX As Long, tpX As Long
    For opX = 1 To ops.count
        Dim opClear As Operation: Set opClear = ops(opX)
        Dim subsClear As SubOperations: Set subsClear = opClear.SubOperations
        If Not (subsClear Is Nothing) Then
            For subX = 1 To subsClear.count
                Dim subClear As SubOperation: Set subClear = subsClear(subX)
                Dim tpsClear As paths: Set tpsClear = subClear.ToolPaths
                If Not (tpsClear Is Nothing) Then
                    For tpX = 1 To tpsClear.count
                        Dim tpClear As Path: Set tpClear = tpsClear(tpX)
                        If Not (tpClear Is Nothing) Then tpClear.OpNo = 0
                    Next tpX
                End If
            Next subX
        End If
    Next opX
    
    ' ===== 阶段3：按排版分组分配 OpNo =====
    App.SetUndoCommandName "排版刀具排序"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    
    Dim baseOpNo As Long: baseOpNo = 1
    Dim sheetId As Long
    
    For sheetId = 1 To sheetCount
        Dim pos As Long: pos = 0
        
        ' 收集属于当前排版的所有 OP
        Dim sheetOpIdxs As New Collection
        For opIdx = 1 To ops.count
            If opSheetMap.Exists(opIdx) Then
                If opSheetMap(opIdx) = sheetId Then
                    sheetOpIdxs.Add opIdx
                End If
            End If
        Next opIdx
        
        If sheetOpIdxs.count = 0 Then GoTo NextSheet2
        
        ' 按用户排序顺序分配
        For sj = 0 To UBound(sortedKeys)
            Dim targetKey As String: targetKey = sortedKeys(sj)
            
            Dim si2 As Long
            For si2 = 1 To sheetOpIdxs.count
                Dim currentOpIdx As Long: currentOpIdx = sheetOpIdxs(si2)
                
                Dim jj As Long
                For jj = 1 To allSubOps.count
                    Dim parts() As String: parts = Split(allSubOps(jj), "|")
                    Dim recOp As Long: recOp = CLng(parts(0))
                    Dim subKey As String: subKey = parts(1)
                    
                    If recOp = currentOpIdx And subKey = targetKey Then
                        Dim newOpNo As Long: newOpNo = baseOpNo + pos
                        pos = pos + 1
                        
                        Dim subParts() As String: subParts = Split(parts(2), ",")
                        Dim nn As Long
                        For nn = 0 To UBound(subParts)
                            Dim recSub As Long: recSub = CLng(subParts(nn))
                            Dim tgtOp As Operation: Set tgtOp = ops(recOp)
                            Dim tgtSub As SubOperation: Set tgtSub = tgtOp.SubOperations(recSub)
                            Dim tps As paths: Set tps = tgtSub.ToolPaths
                            If Not (tps Is Nothing) Then
                                Dim mm As Long
                                For mm = 1 To tps.count
                                    Dim tp As Path: Set tp = tps(mm)
                                    If Not (tp Is Nothing) Then tp.OpNo = newOpNo
                                Next mm
                            End If
                        Next nn
                    End If
                Next jj
            Next si2
        Next sj
        
        baseOpNo = baseOpNo + pos
NextSheet2:
    Next sheetId
    
    ops.OrderAll
    drw.ScreenUpdating = True: drw.Redraw
    Exit Sub
    
ErrHandler3:
    drw.ScreenUpdating = True
    MsgBox "排版刀具排序出错：" & Err.Description & " (0x" & Hex(Err.Number) & ")", vbCritical
End Sub





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

