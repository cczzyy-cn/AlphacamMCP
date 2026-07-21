' ==============================================================================
' CCC功能 - modTest 通用测试模块
' ==============================================================================
' 功能：提供测试入口，用于验证和调试 CCC 功能。
' 后续新增的测试功能统一添加到此模块。
' ==============================================================================
Option Explicit
Option Private Module


' ==============================================================================
' Sub 测试()
' ==============================================================================
' 通用测试入口。由 AlphaCAM 菜单 "CCC功能 > 【测试】" 触发。
' 当前功能：显示选中刀具路径的基本信息（Name、包围盒、中心坐标）。
' ==============================================================================
Sub 测试()
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    
    ' 让用户选择路径
    If Not drw.UserSelectMultiToolPaths("【测试】请选择要查看的刀具路径", 0) Then
        Exit Sub
    End If
    
    Dim tp As Path: Set tp = drw.GetFirstToolPath
    Dim msg As String
    Dim idx As Long: idx = 0
    
    Do While Not (tp Is Nothing)
        If tp.Selected Then
            idx = idx + 1
            msg = msg & "路径 " & idx & ": " & tp.Name & vbCrLf
            msg = msg & "  OpNo=" & tp.OpNo & vbCrLf
            msg = msg & "  X=[" & tp.MinXL & ", " & tp.MaxXL & "]" & vbCrLf
            msg = msg & "  Y=[" & tp.MinYL & ", " & tp.MaxYL & "]" & vbCrLf
            msg = msg & "  centerX=" & ((tp.MinXL + tp.MaxXL) / 2) & vbCrLf
            msg = msg & "  centerY=" & ((tp.MinYL + tp.MaxYL) / 2) & vbCrLf
            msg = msg & vbCrLf
            tp.Selected = False
        End If
        Set tp = tp.GetNext
    Loop
    
    If msg = "" Then msg = "没有选中任何路径！"
    MsgBox msg, vbInformation, "CCC功能测试"
End Sub
