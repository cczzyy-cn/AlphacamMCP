
' ==============================================================================
' CCCFUNC ¡ª modSort ÅÅ°æµ¶¾ßÅÅÐò
' ==============================================================================
Option Explicit
Option Private Module
Sub ÅÅ°æµ¶¾ßÅÅÐò()
    Dim drw As Drawing
    Set g_mapPathToSheet = Nothing
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "No drawing!", vbExclamation, "Sort": Exit Sub
    If drw.Operations Is Nothing Or drw.Operations.count = 0 Then MsgBox "No operations!", vbExclamation, "Sort": Exit Sub
    frmToolSort.Show vbModeless
End Sub
Public Function ScanOperations() As Collection
    Dim result As New Collection
    Dim drw As Drawing, ops As Operations, dict As Object
    Dim i As Long, j As Long
    Dim op As Operation, subs As SubOperations, subop As SubOperation, t As MillTool
    Dim methodName As String, spPos As Integer
    Dim toolDisp As String, key As String, toolD3 As String, keyS As String
    Dim tpIdxS As Long, tpCntS As Long, tpS As Path, tS As MillTool
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then Set ScanOperations = result: Exit Function
    Set ops = drw.Operations
    If ops Is Nothing Then Set ScanOperations = result: Exit Function
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To ops.count
        Set op = ops(i)
        Set subs = op.SubOperations
        If subs Is Nothing Then GoTo NextOp2
        For j = 1 To subs.count
            Set subop = subs(j)
            Set t = subop.Tool
            If Not (t Is Nothing) Then
                methodName = subop.Name
                spPos = InStr(methodName, " "): If spPos > 0 Then methodName = Left(methodName, spPos - 1)

                    If IsNull(t.Name) Or t.Name = "" Then toolDisp = "T" & CStr(t.Number) Else toolDisp = t.Name
                    key = methodName & " T" & CStr(t.Number) & " " & toolDisp
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        Next j
NextOp2:
    Next i
    ' Fallback: scan toolpaths directly when no operations
    If dict.Count = 0 Then

        tpCntS = drw.GetToolPathCount
        If tpCntS > 0 Then
            Set tpS = drw.GetFirstToolPath
            For tpIdxS = 1 To tpCntS
                    Set tS = tpS.GetTool
                    If Not (tS Is Nothing) Then

                        If IsNull(tS.Name) Or tS.Name = "" Then toolD3 = "T" & CStr(tS.Number) Else toolD3 = tS.Name
                        keyS = "T" & CStr(tS.Number) & " " & toolD3
                        If Not dict.Exists(keyS) Then dict.Add keyS, True
                    End If
                End If
            Next tpIdxS
        End If
    End If
    Dim keysArr: keysArr = dict.Keys
    Dim kk As Long
    For kk = 0 To UBound(keysArr)
        result.Add keysArr(kk)
    Next kk
    Set ScanOperations = result
End Function
Public Sub ApplySortToDrawing(ByRef sortedKeys() As String)
    On Error GoTo ErrHandler3
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub
    Dim drw As Drawing, ops As Operations
    Dim opIdx As Long, s As Long, si As Long, sj As Long, mi As Long
    Dim mSheet As Long, pos As Long, lastSh As Long, spInt As Integer
    Dim firstTpName As String, lookupName As String
    Dim ta As Path, tc As Collection, c2 As Collection, cA As Collection, nc As Collection, ncA As Collection
    Dim siN As Long, toolD2 As String, toolD4 As String, tpIdxA As Long, tpA As Path, tA As MillTool, stD As Object, nd As Object, sp As paths, subF As SubOperation, subL As SubOperation, subM As SubOperation, subN As SubOperation, tpF As paths, tpL As paths, tpM As Path, tpN As paths, tpCntA As Long, tF As Path, tL As Path, tM As MillTool, tN2 As Path, tpc As paths, tpc2 As Path, tpsM As paths, sbc As SubOperation, sc As SubOperations, oc As Operation, opN As Operation, opM As Operation, sbN As SubOperations, sbM As SubOperations, tk As String, ck As String, tky As String, ck2 As String, tkA As String, ckA As String, sheetCount As Long, shNm, ni As NestInformation
    If g_mapPathToSheet Is Nothing Then
        Set g_mapPathToSheet = CreateObject("Scripting.Dictionary")
        Set ni = drw.GetNestInformation()
        sheetCount = ni.Sheets.count
        If sheetCount = 0 Then sheetCount = 1
        ReDim shNm(1 To sheetCount)
        For s = 1 To sheetCount
            Set nd = CreateObject("Scripting.Dictionary")
            Set sp = ni.Sheets(s).paths
            For mi = 1 To sp.count
                If Not nd.Exists(sp(mi).Name) Then nd.Add sp(mi).Name, True
            Next mi
            Set shNm(s) = nd
        Next s
        For opIdx = 1 To ops.count
            Set opN = ops(opIdx)
            Set sbN = opN.SubOperations
            mSheet = 0: firstTpName = ""
            If Not (sbN Is Nothing) Then
                If sbN.count > 0 Then
                    Set subF = sbN(1)
                    Set tpF = subF.ToolPaths
                    If Not (tpF Is Nothing) Then
                        If tpF.count > 0 Then
                            Set tF = tpF(1)
                            If Not (tF Is Nothing) Then firstTpName = tF.Name
                        End If
                    End If
                End If
            End If
            If Not (sbN Is Nothing) Then
                For siN = 1 To sbN.count
                    Set subN = sbN(siN)
                    Set tpN = subN.ToolPaths
                    If Not (tpN Is Nothing) Then
                        For mi = 1 To tpN.count
                            Set tN2 = tpN(mi)
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
            If mSheet = 0 And lastSh > 0 Then mSheet = lastSh
            If mSheet = 0 Then mSheet = 1
            If firstTpName <> "" Then
                If Not g_mapPathToSheet.Exists(firstTpName) Then
                    g_mapPathToSheet.Add firstTpName, mSheet
                End If
            End If
        Next opIdx
    End If
    Set stD = CreateObject("Scripting.Dictionary")
    sheetId = 1: mn = "": tk = "": ck = ""
    For opIdx = 1 To ops.count
        Set opM = ops(opIdx)
        Set sbM = opM.SubOperations
        If Not (sbM Is Nothing) Then
            lookupName = ""
            If sbM.count > 0 Then
                Set subL = sbM(1)
                Set tpL = subL.ToolPaths
                If Not (tpL Is Nothing) Then
                    If tpL.count > 0 Then
                        Set tL = tpL(1)
                        If Not (tL Is Nothing) Then lookupName = tL.Name
                    End If
                End If
            End If
            sheetId = 1
            If lookupName <> "" And g_mapPathToSheet.Exists(lookupName) Then
                sheetId = g_mapPathToSheet(lookupName)
            End If
            For si = 1 To sbM.count
                Set subM = sbM(si)
                Set tM = subM.Tool
                If Not (tM Is Nothing) Then
                    mn = subM.Name
                    spInt = InStr(mn, "  ")
                    If spInt > 0 Then mn = Left(mn, spInt - 1) Else: spInt = InStr(mn, " "): If spInt > 0 Then mn = Left(mn, spInt - 1)
                    If IsNull(tM.Name) Or tM.Name = "" Then toolD2 = "T" & CStr(tM.Number) Else toolD2 = tM.Name
                    tk = mn & " T" & CStr(tM.Number) & " " & toolD2
                    ck = CStr(sheetId) & "|" & tk
                    Set tpsM = subM.ToolPaths
                    If Not (tpsM Is Nothing) Then
                        For mi = 1 To tpsM.count
                            Set tpM = tpsM(mi)
                            If Not (tpM Is Nothing) Then
                                If Not stD.Exists(ck) Then
                                    Set nc = New Collection
                                    stD.Add ck, nc
                                End If
                                Set c2 = stD(ck)
                                c2.Add tpM
                            End If
                        Next mi
                    End If
                End If
            Next si
        End If
    Next opIdx
    ' Fallback: direct toolpath sort when no operations
    If stD.Count = 0 Then
        tpCntA = drw.GetToolPathCount
        If tpCntA > 0 Then
            Set tpA = drw.GetFirstToolPath
            For tpIdxA = 1 To tpCntA
                If Not (tpA Is Nothing) Then
                    Set tA = tpA.GetTool
                    If Not (tA Is Nothing) Then
                        If IsNull(tA.Name) Or tA.Name = "" Then toolD4 = "T" & CStr(tA.Number) Else toolD4 = tA.Name
                        tkA = "T" & CStr(tA.Number) & " " & toolD4
                        ckA = "1|" & tkA
                        If Not stD.Exists(ckA) Then
                            Set ncA = New Collection
                            stD.Add ckA, ncA
                        End If
                        Set cA = stD(ckA)
                        cA.Add tpA
                    End If
                    Set tpA = tpA.GetNext
                End If
            Next tpIdxA
        End If
    End If
    If stD.count = 0 Then Exit Sub
    Dim ox As Long, sx As Long, tx As Long
    For ox = 1 To ops.count
        Set oc = ops(ox)
        Set sc = oc.SubOperations
        If Not (sc Is Nothing) Then
            For sx = 1 To sc.count
                Set sbc = sc(sx)
                Set tpc = sbc.ToolPaths
                If Not (tpc Is Nothing) Then
                    For tx = 1 To tpc.count
                        Set tpc2 = tpc(tx)
                        If Not (tpc2 Is Nothing) Then tpc2.OpNo = 0
                    Next tx
                End If
            Next sx
        End If
    Next ox
    App.SetUndoCommandName "Sort"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    For si = 1 To sheetCount
        pos = 1
        For sj = 0 To UBound(sortedKeys)
            tky = sortedKeys(sj)
            ck2 = CStr(si) & "|" & tky
            If stD.Exists(ck2) Then
                Set tc = stD(ck2)
                For mi = 1 To tc.count
                    Set ta = Nothing
                    On Error Resume Next
                    Set ta = tc(mi)
                    On Error GoTo 0
                    If Not (ta Is Nothing) Then ta.OpNo = si * 1000 + pos
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
    MsgBox "Sort error: " & Err.Description, vbCritical
End Sub
