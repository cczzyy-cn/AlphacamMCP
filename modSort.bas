
' ==============================================================================
' CCCFUNC ¡ª modSort ÅÅ°æµ¶¾ßÅÅÐò
' ==============================================================================
Option Explicit
Option Private Module

Sub ÅÅ°æµ¶¾ßÅÅÐò()
    Set g_mapPathToSheet = Nothing
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "No drawing!", vbExclamation, "Sort": Exit Sub
    If drw.Operations Is Nothing Or drw.Operations.count = 0 Then MsgBox "No operations!", vbExclamation, "Sort": Exit Sub
    frmToolSort.Show vbModeless
End Sub

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
                Dim key As String: key = methodName & " T" & CStr(t.Number) & " " & IIf(t.Name <> "", t.Name, "T" & CStr(t.Number))
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        Next j
NextOp2:
    Next i
    ' Fallback: scan toolpaths directly when no operations
    If dict.Count = 0 Then
        Dim tpIdxS As Long
        Dim tpCntS As Long: tpCntS = drw.GetToolPathCount
        If tpCntS > 0 Then
            Dim tpS As Path: Set tpS = drw.GetFirstToolPath
            For tpIdxS = 1 To tpCntS
                If Not (tpS Is Nothing) Then
                    Dim tS As MillTool: Set tS = tpS.GetTool
                    If Not (tS Is Nothing) Then
                        Dim keyS As String: keyS = "T" & CStr(tS.Number) & " " & IIf(tS.Name <> "", tS.Name, "T" & CStr(tS.Number))
                        If Not dict.Exists(keyS) Then dict.Add keyS, True
                    End If
                    Set tpS = tpS.GetNext
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
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    Dim ops As Operations: Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub
    Dim opIdx As Long, s As Long, si As Long, sj As Long, mi As Long
    Dim mSheet As Long, pos As Long, lastSh As Long, spInt As Integer
    Dim firstTpName As String, lookupName As String
    If g_mapPathToSheet Is Nothing Then
        Set g_mapPathToSheet = CreateObject("Scripting.Dictionary")
        Dim ni As NestInformation: Set ni = drw.GetNestInformation()
        Dim sheetCount As Long: sheetCount = ni.Sheets.count
        If sheetCount = 0 Then sheetCount = 1
        Dim shNm() As Object: ReDim shNm(1 To sheetCount)
        For s = 1 To sheetCount
            Dim nd As Object: Set nd = CreateObject("Scripting.Dictionary")
            Dim sp As paths: Set sp = ni.Sheets(s).paths
            For mi = 1 To sp.count
                If Not nd.Exists(sp(mi).Name) Then nd.Add sp(mi).Name, True
            Next mi
            Set shNm(s) = nd
        Next s
        For opIdx = 1 To ops.count
            Dim opN As Operation: Set opN = ops(opIdx)
            Dim sbN As SubOperations: Set sbN = opN.SubOperations
            mSheet = 0: firstTpName = ""
            If Not (sbN Is Nothing) Then
                If sbN.count > 0 Then
                    Dim subF As SubOperation: Set subF = sbN(1)
                    Dim tpF As paths: Set tpF = subF.ToolPaths
                    If Not (tpF Is Nothing) Then
                        If tpF.count > 0 Then
                            Dim tF As Path: Set tF = tpF(1)
                            If Not (tF Is Nothing) Then firstTpName = tF.Name
                        End If
                    End If
                End If
            End If
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
            If mSheet = 0 And lastSh > 0 Then mSheet = lastSh
            If mSheet = 0 Then mSheet = 1
            If firstTpName <> "" Then
                If Not g_mapPathToSheet.Exists(firstTpName) Then
                    g_mapPathToSheet.Add firstTpName, mSheet
                End If
            End If
        Next opIdx
    End If
    Dim stD As Object: Set stD = CreateObject("Scripting.Dictionary")
    Dim sheetId As Long, mn As String, tk As String, ck As String
    Dim tM As MillTool
    For opIdx = 1 To ops.count
        Dim opM As Operation: Set opM = ops(opIdx)
        Dim sbM As SubOperations: Set sbM = opM.SubOperations
        If Not (sbM Is Nothing) Then
            lookupName = ""
            If sbM.count > 0 Then
                Dim subL As SubOperation: Set subL = sbM(1)
                Dim tpL As paths: Set tpL = subL.ToolPaths
                If Not (tpL Is Nothing) Then
                    If tpL.count > 0 Then
                        Dim tL As Path: Set tL = tpL(1)
                        If Not (tL Is Nothing) Then lookupName = tL.Name
                    End If
                End If
            End If
            sheetId = 1
            If lookupName <> "" And g_mapPathToSheet.Exists(lookupName) Then
                sheetId = g_mapPathToSheet(lookupName)
            End If
            For si = 1 To sbM.count
                Dim subM As SubOperation: Set subM = sbM(si)
                Set tM = subM.Tool
                If Not (tM Is Nothing) Then
                    mn = subM.Name
                    spInt = InStr(mn, "  ")
                    If spInt > 0 Then mn = Left(mn, spInt - 1) Else: spInt = InStr(mn, " "): If spInt > 0 Then mn = Left(mn, spInt - 1)
                    tk = mn & " T" & CStr(tM.Number) & " " & IIf(tM.Name <> "", tM.Name, "T" & CStr(tM.Number))
                    ck = CStr(sheetId) & "|" & tk
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
    ' Fallback: direct toolpath sort when no operations
    If stD.Count = 0 Then
        Dim tpIdxA As Long
        Dim tpCntA As Long: tpCntA = drw.GetToolPathCount
        If tpCntA > 0 Then
            Dim tpA As Path: Set tpA = drw.GetFirstToolPath
            For tpIdxA = 1 To tpCntA
                If Not (tpA Is Nothing) Then
                    Dim tA As MillTool: Set tA = tpA.GetTool
                    If Not (tA Is Nothing) Then
                        Dim tkA As String: tkA = "T" & CStr(tA.Number) & " " & IIf(tA.Name <> "", tA.Name, "T" & CStr(tA.Number))
                        Dim ckA As String: ckA = "1|" & tkA
                        If Not stD.Exists(ckA) Then
                            Dim ncA As Collection: Set ncA = New Collection
                            stD.Add ckA, ncA
                        End If
                        Dim cA As Collection: Set cA = stD(ckA)
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
    App.SetUndoCommandName "Sort"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    For si = 1 To sheetCount
        pos = 1
        Dim tky As String, ck2 As String, tc As Collection, ta As Path
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
