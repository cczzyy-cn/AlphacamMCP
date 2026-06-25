' ==============================================================================
' CCC功能 — modOffset 全排版刀具偏移
' ==============================================================================
Option Explicit
Option Private Module

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
                    ' 匹配逻辑：精确名 → 包含名 → T 号
                    Dim isMatch As Boolean: isMatch = False
                    If t.Name = selectedTool Then
                        isMatch = True
                    ElseIf InStr(1, t.Name, selectedTool, vbTextCompare) > 0 Then
                        isMatch = True
                    ElseIf Not foundFirst And InStr(1, CStr(t.Number), selectedTool, vbTextCompare) > 0 Then
                        isMatch = True
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
            Next j
        End If
    Next i
    End If  ' End If ops.count > 0
    drw.ScreenUpdating = True: drw.Redraw
    MsgBox "偏移完成！已处理 " & count & " 条刀具路径。" & vbCrLf & "刀具: " & selectedTool & vbCrLf & "偏移量: X=" & xOff & "  Y=" & yOff & "  Z=" & zOff, vbInformation, "全排版刀具偏移"
    Exit Sub
ErrHandler2:
    drw.ScreenUpdating = True
    MsgBox "偏移出错：" & Err.Description, vbCritical
End Sub
