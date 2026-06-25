
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
                Dim toolDisp As String
                    If IsNull(t.Name) Or t.Name = "" Then toolDisp = "T" & CStr(t.Number) Else toolDisp = t.Name
                    Dim key As String: key = methodName & " T" & CStr(t.Number) & " " & toolDisp
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
                        Dim toolD3 As String
                        If IsNull(tS.Name) Or tS.Name = "" Then toolD3 = "T" & CStr(tS.Number) Else toolD3 = tS.Name
                        Dim keyS As String: keyS = "T" & CStr(tS.Number) & " " & toolD3
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
        Dim tx As Long
        Dim tpsM As paths
        Dim tpc2 As Path
        Dim tpc As paths
        Dim tpN As paths
        Dim tpM As Path
        Dim tpL As paths
        Dim tpIdxA As Long
        Dim tpF As paths
        Dim tpCntA As Long
        Dim tpA As Path
        Dim toolD4 As String
        Dim toolD2 As String
        Dim tky As String
        Dim tkA As String
        Dim tk As String
        Dim tN2 As Path
        Dim tM As MillTool
        Dim tL As Path
        Dim tF As Path
    If drw Is Nothing Then Exit Sub
    Dim ops As Operations: Set ops = drw.Operations
    If ops Is Nothing Then Exit Sub
    Dim opIdx As Long, s As Long, si As Long, sj As Long, mi As Long
    Dim tc As Collection, ta As Path, ncTemp As Collection, ncATemp As Collection, c2 As Collection, cA As Collection
    Dim mSheet As Long, pos As Long, lastSh As Long, spInt As Integer
    Dim firstTpName As String, lookupName As String
    If g_mapPathToSheet Is Nothing Then
        Set g_mapPathToSheet = CreateObject("Scripting.Dictionary")


        If sheetCount = 0 Then sheetCount = 1

        For s = 1 To sheetCount


            For mi = 1 To sp.count
                If Not nd.Exists(sp(mi).Name) Then nd.Add sp(mi).Name, True
            Next mi
            Set shNm(s) = nd
        Next s
        For opIdx = 1 To ops.count


            mSheet = 0: firstTpName = ""
            If Not (sbN Is Nothing) Then
                If sbN.count > 0 Then


                    If Not (tpF Is Nothing) Then
                        If tpF.count > 0 Then

                            If Not (tF Is Nothing) Then firstTpName = tF.Name
                        End If
                    End If
                End If
            End If
            If Not (sbN Is Nothing) Then

                For siN = 1 To sbN.count


                    If Not (tpN Is Nothing) Then
                        For mi = 1 To tpN.count

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



    For opIdx = 1 To ops.count


        If Not (sbM Is Nothing) Then
            lookupName = ""
            If sbM.count > 0 Then


                If Not (tpL Is Nothing) Then
                    If tpL.count > 0 Then

                        If Not (tL Is Nothing) Then lookupName = tL.Name
                    End If
                End If
            End If
            sheetId = 1
            If lookupName <> "" And g_mapPathToSheet.Exists(lookupName) Then
                sheetId = g_mapPathToSheet(lookupName)
            End If
            For si = 1 To sbM.count

                Set tM = subM.Tool
                If Not (tM Is Nothing) Then
                    mn = subM.Name
                    spInt = InStr(mn, "  ")
                    If spInt > 0 Then mn = Left(mn, spInt - 1) Else: spInt = InStr(mn, " "): If spInt > 0 Then mn = Left(mn, spInt - 1)

                    If IsNull(tM.Name) Or tM.Name = "" Then toolD2 = "T" & CStr(tM.Number) Else toolD2 = tM.Name
                    tk = mn & " T" & CStr(tM.Number) & " " & toolD2
                    ck = CStr(sheetId) & "|" & tk

                    If Not (tpsM Is Nothing) Then
                        For mi = 1 To tpsM.count

                            If Not (tpM Is Nothing) Then
                                If Not stD.Exists(ck) Then

                                    stD.Add ck, ncTemp
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


        If tpCntA > 0 Then

            For tpIdxA = 1 To tpCntA
                If Not (tpA Is Nothing) Then

                    If Not (tA Is Nothing) Then

                        If IsNull(tA.Name) Or tA.Name = "" Then toolD4 = "T" & CStr(tA.Number) Else toolD4 = tA.Name


                        If Not stD.Exists(ckA) Then
                            Set ncATemp = New Collection
                            stD.Add ckA, ncATemp
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


        If Not (sc Is Nothing) Then
            For sx = 1 To sc.count


                If Not (tpc Is Nothing) Then
                    For tx = 1 To tpc.count

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


            If stD.Exists(ck2) Then
                Set tc = stD(ck2)
                For mi = 1 To tc.count
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
