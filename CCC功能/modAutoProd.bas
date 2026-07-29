Option Explicit

' ==============================================================================
' 自动化生产排版 — 一键导入CSV → 批量生产 → 排版 → NC输出
' ==============================================================================
' 依赖: CDM.arb 已加载（BatchImport 模块 + g_Make_Master 函数）
' 用法: 点击 CCC功能 → 自动化生产排版 → 输入CSV文件名
' ==============================================================================

Public Sub 自动化生产排版()
    '
    Dim sCSVPath As String
    Dim sJobName As String
    Dim sFolder  As String
    Dim sTemp    As String
    Dim lngOrderID As Long
    Dim rst      As ADODB.Recordset
    
    On Error GoTo EH
    
    ' ── 1. 输入 CSV 文件名 ──
    sCSVPath = InputBox( _
        "请输入 CSV 文件名（含 .csv 扩展名）" & vbCrLf & vbCrLf & _
        "将从以下目录读取：" & vbCrLf & _
        "C:\Users\C\Desktop\2026优化表\", _
        "自动化生产排版", _
        "7-10中林SPC婷兰灰.csv")
    
    ' 点击取消
    If sCSVPath = "" Then Exit Sub
    
    ' 如果用户只输入了文件名，补全路径
    sFolder = "C:\Users\C\Desktop\2026优化表\"
    If InStr(sCSVPath, "\") = 0 And InStr(sCSVPath, "/") = 0 Then
        sCSVPath = sFolder & sCSVPath
    End If
    
    ' ── 2. 检查文件是否存在 ──
    If Dir(sCSVPath) = "" Then
        MsgBox "文件不存在:" & vbCrLf & sCSVPath, vbExclamation, "自动化生产排版"
        GoTo CleanUp
    End If
    
    ' 从 CSV 文件名获取订单名（去掉 .csv）
    sTemp = sCSVPath
    Do While InStr(sTemp, "\") > 0
        sTemp = Mid$(sTemp, InStr(sTemp, "\") + 1)
    Loop
    Do While InStr(sTemp, "/") > 0
        sTemp = Mid$(sTemp, InStr(sTemp, "/") + 1)
    Loop
    If LCase(Right$(sTemp, 4)) = ".csv" Then
        sJobName = Left$(sTemp, Len(sTemp) - 4)
    Else
        sJobName = sTemp
    End If
    
    ' ── 3. 检查 BatchImport 是否已安装 ──
    On Error Resume Next
    Call BatchImport.Run(sCSVPath, sJobName)
    If Err.Number <> 0 Then
        Dim sErrMsg As String
        sErrMsg = Err.Description
        On Error GoTo EH
        If InStr(sErrMsg, "未找到") Then
            MsgBox "BatchImport 模块未安装！" & vbCrLf & vbCrLf & _
                   "请先在 VBA 编辑器中导入:" & vbCrLf & _
                   "CDM功能/BatchImport.bas", vbCritical, "自动化生产排版"
        Else
            MsgBox "CSV 导入失败:" & vbCrLf & sErrMsg, vbCritical, "自动化生产排版"
        End If
        GoTo CleanUp
    End If
    On Error GoTo EH
    
    ' ── 4. 查找刚创建的订单 ID ──
    If Not gbln_ConnectToDB() Then
        MsgBox "无法连接 CDM 数据库", vbCritical, "自动化生产排版"
        GoTo CleanUp
    End If
    
    Set rst = gdb_CDM.Execute("SELECT OrderID FROM AD_ORDERS " & _
        "WHERE JobName='" & gs_FixSQL(sJobName) & "' " & _
        "ORDER BY OrderID DESC")
    
    If rst.BOF And rst.EOF Then
        MsgBox "未找到订单，CSV 导入可能失败", vbExclamation, "自动化生产排版"
        rst.Close
        GoTo CleanUp
    End If
    
    lngOrderID = rst.Fields("OrderID").Value
    rst.Close
    
    ' ── 5. 执行批量生产 + 排版 ──
    ' 锁定屏幕加速
    App.Frame.ProjectBarUpdating = False
    App.DisableUndo = True
    
    ' 先清除上次完成标记
    SaveSetting "LICOM AlphaDOOR", "Nest Parameters", "Nest Completed", 0
    
    ' 调用 CDM 加工引擎
    Call g_Make_Master(CStr(lngOrderID))
    
    ' 检查加工结果
    Dim bSuccess As Boolean
    bSuccess = CBool(GetSetting("LICOM AlphaDOOR", "Nest Parameters", "Nest Completed", 0))
    
    ' ── 6. 完成 ──
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
    GoTo CleanUp
    
EH:
    MsgBox "错误: " & Err.Description & vbCrLf & _
           "位置: " & Erl, vbCritical, "自动化生产排版"
    
CleanUp:
    Set rst = Nothing
    App.Frame.ProjectBarUpdating = True
    App.DisableUndo = False
End Sub
