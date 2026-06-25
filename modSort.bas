
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
    Dim dict As Object, drw As Drawing, i As Long, j As Long, key As String, keyS As String, kk As Long, methodName As String, op As Operation, ops As Operations, result As New Collection, spPos As Integer, subop As SubOperation, subs As SubOperations, t As MillTool, tS As MillTool, tpCntS As Long, tpIdxS As Long, tpS As Path

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

                spPos = InStr(methodName, "  ")
                If spPos > 0 Then methodName = Left(methodName, spPos - 1) Else: spPos = InStr(methodName, " "): If spPos > 0 Then methodName = Left(methodName, spPos - 1)
                key = methodName & " T" & CStr(t.Number) & " " & IIf(t.Name <> "", t.Name, "T" & CStr(t.Number))
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
                If Not (tpS Is Nothing) Then
                    Set tS = tpS.GetTool
                    If Not (tS Is Nothing) Then
                        keyS = "T" & CStr(tS.Number) & " " & IIf(tS.Name <> "", tS.Name, "T" & CStr(tS.Number))
                        If Not dict.Exists(keyS) Then dict.Add keyS, True
                    End If
                    Set tpS = tpS.GetNext
                End If
            Next tpIdxS
        End If
    End If
    Dim keysArr: keysArr = dict.Keys

    For kk = 0 To UBound(keysArr)
        result.Add keysArr(kk)
    Next kk
    Set ScanOperations = result
End Function

Public Sub ApplySortToDrawing(ByRef sortedKeys() As String)
    Dim c2 As Collection, cA As Collection, ck As String, ck2 As String, ckA As String
    Dim drw As Drawing, firstTpName As String, lastSh As Long, lookupName As String
    Dim mSheet As Long, mi As Long, mn As String, nc As Collection, ncA As Collection
    Dim nd As Object, ni As NestInformation, oc As Operation, opIdx As Long
    Dim opM As Operation, opN As Operation, ops As Operations, ox As Long, pos As Long
    Dim s As Long, sbM As SubOperations, sbN As SubOperations, sbc As SubOperation
    Dim sc As SubOperations, shNm() As Object, sheetCount As Long, sheetId As Long
    Dim si As Long, siN As Long, sj As Long, sp As paths, spInt As Integer
    Dim stD As Object, subF As SubOperation, subL As SubOperation, subM As SubOperation
    Dim subN As SubOperation, sx As Long, tA As MillTool, tF As Path, tL As Path
    Dim tM As MillTool, tN2 As Path, ta As Path, tc As Collection, tk As String
    Dim tkA As String, tky As String, tpA As Path, tpCntA As Long, tpF As paths
    Dim tpIdxA As Long, tpL As paths, tpM As Path, tpN As paths, tpc As paths
    Dim tpc2 As Path, tpsM As paths, tx As Long
    On Error GoTo ErrHandler3
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub

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
                    tk = mn & " T" & CStr(tM.Number) & " " & IIf(tM.Name <> "", tM.Name, "T" & CStr(tM.Number))
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
                        tkA = "T" & CStr(tA.Number) & " " & IIf(tA.Name <> "", tA.Name, "T" & CStr(tA.Number))
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
                    Set ta = tc(mi)
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
