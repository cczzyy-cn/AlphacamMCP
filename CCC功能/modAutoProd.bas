Option Explicit

' ==============================================================================
' 自动化生产排版 — 一键导入CSV → 批量生产 → 排版 → NC输出
' ==============================================================================
' 说明: 通过 Application.Run 跨工程调用 CDM 模块
' 依赖: CDM 工程中已安装 BatchImport 模块
' 用法: 点击 CCC功能 → 自动化生产排版 → 输入CSV文件名
' ==============================================================================

Public Sub 自动化生产排版()
    '
    Dim sCSVPath As String
    Dim sJobName As String
    Dim sFolder  As String
    Dim sTemp    As String
    
    On Error GoTo EH
    
    ' ── 1. 输入 CSV 文件名 ──
    sCSVPath = InputBox( _
        "请输入 CSV 文件名（含 .csv 扩展名）" & vbCrLf & vbCrLf & _
        "将从以下目录读取：" & vbCrLf & _
        "C:\Users\C\Desktop\2026优化表\", _
        "自动化生产排版", _
        "7-10中林SPC婷兰灰.csv")
    
    If sCSVPath = "" Then Exit Sub
    
    ' 补全路径
    sFolder = "C:\Users\C\Desktop\2026优化表\"
    If InStr(sCSVPath, "\") = 0 And InStr(sCSVPath, "/") = 0 Then
        sCSVPath = sFolder & sCSVPath
    End If
    
    ' ── 2. 检查文件是否存在 ──
    If Dir(sCSVPath) = "" Then
        MsgBox "文件不存在:" & vbCrLf & sCSVPath, vbExclamation, "自动化生产排版"
        Exit Sub
    End If
    
    ' 从文件名取订单名
    sTemp = sCSVPath
    Do While InStr(sTemp, "\") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "\") + 1): Loop
    Do While InStr(sTemp, "/") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "/") + 1): Loop
    If LCase(Right$(sTemp, 4)) = ".csv" Then
        sJobName = Left$(sTemp, Len(sTemp) - 4)
    Else
        sJobName = sTemp
    End If
    
    ' ── 3. 调用 CDM.BatchImport.Run 导入 CSV ──
    App.Frame.ShowProgressBox "自动化生产排版", "正在导入 CSV ..."
    DoEvents
    
    On Error Resume Next
    Application.Run "CDM.BatchImport.Run", sCSVPath, sJobName
    If Err.Number <> 0 Then
        Dim sErr As String: sErr = Err.Description
        On Error GoTo EH
        App.Frame.CloseProgressBox
        If InStr(sErr, "Object required") Or InStr(sErr, "未找到") Then
            MsgBox "BatchImport 模块未安装！" & vbCrLf & vbCrLf & _
                   "请先在 CDM 工程的 VBA 编辑器中导入:" & vbCrLf & _
                   "CDM功能/BatchImport.bas", vbCritical, "自动化生产排版"
        Else
            MsgBox "CSV 导入失败:" & vbCrLf & sErr, vbCritical, "自动化生产排版"
        End If
        Exit Sub
    End If
    On Error GoTo EH
    
    ' ── 4. 查找 OrderID ──
    '    通过注册表获取（BatchImport 会保存）
    Dim lngOrderID As Long
    lngOrderID = CLng(GetSetting("CCC", "AutoProd", "LastOrderID", "0"))
    
    If lngOrderID <= 0 Then
        App.Frame.CloseProgressBox
        MsgBox "无法获取订单 ID，请检查导入是否成功", vbExclamation, "自动化生产排版"
        Exit Sub
    End If
    
    ' ── 5. 调用 CDM.g_Make_Master 批量生产 ──
    App.Frame.SetProgressText "正在执行批量生产 + 排版 ..."
    DoEvents
    
    Application.Run "CDM.g_Make_Master", CStr(lngOrderID)
    
    App.Frame.CloseProgressBox
    
    ' ── 6. 完成 ──
    Dim bSuccess As Boolean
    bSuccess = CBool(GetSetting("LICOM AlphaDOOR", "Nest Parameters", "Nest Completed", 0))
    
    If bSuccess Then
        MsgBox "自动化生产排版完成！" & vbCrLf & vbCrLf & _
               "订单: " & sJobName & vbCrLf & _
               "后续可在 AlphaCAM 中查看结果并输出 NC", vbInformation, "自动化生产排版"
    Else
        MsgBox "生产排版可能未完全成功，请检查 AlphaCAM 中的结果。" & vbCrLf & vbCrLf & _
               "常见原因:" & vbCrLf & _
               "  - 门型刀路未配置 (AD_DOOR_PATHS)" & vbCrLf & _
               "  - 材料参数不正确" & vbCrLf & _
               "  - 后处理器未选择", vbExclamation, "自动化生产排版"
    End If
    Exit Sub
    
EH:
    App.Frame.CloseProgressBox
    MsgBox "错误: " & Err.Description & vbCrLf & _
           "位置: " & Erl, vbCritical, "自动化生产排版"
End Sub
