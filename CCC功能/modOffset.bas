' ==============================================================================
' CCC功能 — modOffset 全排版刀具偏移
' ==============================================================================
Option Explicit
Option Private Module

Sub 全排版刀具偏移()
    frmToolOffset.Show vbModeless
End Sub

Public Sub ApplyToolOffset(ByVal selectedTool As String, ByVal xOff As Double, ByVal yOff As Double, ByVal zOff As Double, Optional ByVal processName As String = "")
    On Error GoTo ErrHandler2
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    ' Defensive: if selectedTool starts with -T, strip leading -
    Dim selTool As String: selTool = selectedTool
    If Left(selTool, 2) = "-T" Then selTool = Mid(selTool, 2)
    App.SetUndoCommandName "全排版刀具偏移"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    Dim count As Long
    Dim ops As Operations: Set ops = drw.Operations
    Dim i As Long, j As Long
    ' 匹配策略：先精确匹配 t.Name，再尝试包含匹配，最后匹配 t.Number
    Dim foundFirst As Boolean: foundFirst = False
    Dim firstTNum As Long: firstTNum = 0
    For i = 1 To ops.count
        Dim op As Operation: Set op = ops(i)
        Dim subs As SubOperations: Set subs = op.SubOperations
        If Not (subs Is Nothing) Then
            For j = 1 To subs.count
                Dim subop As SubOperation: Set subop = subs(j)
                Dim t As MillTool: Set t = subop.Tool
                If Not (t Is Nothing) Then
                    ' 匹配逻辑：先检查加工方式（如有），再匹配刀具
                    Dim isMatch As Boolean: isMatch = False
                    Dim procMatch As Boolean: procMatch = False
                    ' 如果传入了加工方式名，先检查 subop 的加工方式是否匹配
                    If processName <> "" Then
                        Dim rawName As String: rawName = subop.Name
                        Dim spPos As Integer: spPos = InStr(rawName, "  ")
                        If spPos > 0 Then
                            procMatch = (Trim(Left(rawName, spPos - 1)) = processName)
                        Else
                            spPos = InStr(rawName, " ")
                            If spPos > 0 Then procMatch = (Trim(Left(rawName, spPos - 1)) = processName) _
                                Else procMatch = (rawName = processName)
                        End If
                    Else
                        procMatch = True  ' 未指定加工方式时，不限制
                    End If
                    If procMatch Then
                        If t.Name = selTool Then
                            isMatch = True
                        ElseIf InStr(1, t.Name, selTool, vbTextCompare) > 0 Then
                            isMatch = True
                        ElseIf Not foundFirst And CStr(t.Number) = selTool Then
                            isMatch = True
                        ElseIf Left(selTool, 1) = "T" Then
                            Dim tNumVal As Long: tNumVal = Val(Mid(selTool, 2))
                            If tNumVal > 0 And t.Number = tNumVal Then isMatch = True
                        End If
                    End If
                    If isMatch Then
                        If Not foundFirst Then foundFirst = True: firstTNum = t.Number
                        ' 如果匹配的 T 号与首次匹配的不同则跳过（多把同名刀时取第一把）
                        If t.Number = firstTNum Then
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
            End If
            Next j
        End If
    Next i
        ' Fallback: direct toolpath iteration when no operations（仅匹配刀具，无法检查加工方式）
    If count = 0 Then
        Dim tpIdx2 As Long
        Dim tpCnt2 As Long: tpCnt2 = drw.GetToolPathCount
        If tpCnt2 > 0 Then
            Dim tpP2 As Path: Set tpP2 = drw.GetFirstToolPath
            For tpIdx2 = 1 To tpCnt2
                If Not (tpP2 Is Nothing) Then
                    Dim t2 As MillTool: Set t2 = tpP2.GetTool
                    If Not (t2 Is Nothing) Then
                        If t2.Name = selTool Or (selTool <> "" And (InStr(1, t2.Name, selTool, vbTextCompare) > 0 Or CStr(t2.Number) = selTool)) Then
                            tpP2.MoveG xOff, yOff, zOff: count = count + 1
                        End If
                    End If
                    Set tpP2 = tpP2.GetNext
                End If
            Next tpIdx2
        End If
    End If

drw.ScreenUpdating = True: drw.Redraw
    MsgBox "偏移完成！已处理 " & count & " 条刀具路径。" & vbCrLf & "刀具: " & selectedTool & vbCrLf & "偏移量: X=" & xOff & "  Y=" & yOff & "  Z=" & zOff, vbInformation, "全排版刀具偏移"
    Exit Sub
ErrHandler2:
    drw.ScreenUpdating = True
    MsgBox "偏移出错：" & Err.Description, vbCritical
End Sub
