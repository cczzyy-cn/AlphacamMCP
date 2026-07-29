Attribute VB_Name = "modAutoProd"
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
    Dim FSO      As New Scripting.FileSystemObject
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
    If Not FSO.FileExists(sCSVPath) Then
        MsgBox "文件不存在:" & vbCrLf & sCSVPath, vbExclamation, "自动化生产排版"
        GoTo CleanUp
    End If
    
    ' 从 CSV 文件名获取订单名（去掉 .csv）
    sJobName = FSO.GetBaseName(sCSVPath)
    
    ' ── 3. 检查 BatchImport 模块是否已安装 ──
    On Error Resume Next
    Call BatchImport.Run(sCSVPath, sJobName)
    If Err.Number <> 0 Then
        On Error GoTo EH
        MsgBox "BatchImport 模块未安装或出错。" & vbCrLf & _
               "请先通过 AlphaCAM VBA 编辑器导入 CDM功能/BatchImport.bas", vbCritical
        GoTo CleanUp
    End If
    On Error GoTo EH
    
    ' ── 4. 查找刚创建的订单 ID ──
    If Not gbln_ConnectToDB() Then
        MsgBox "无法连接 CDM 数据库", vbCritical
        GoTo CleanUp
    End If
    
    Set rst = gdb_CDM.Execute("SELECT OrderID FROM AD_ORDERS " & _
        "WHERE JobName='" & gs_FixSQL(sJobName) & "' " & _
        "ORDER BY OrderID DESC")
    
    If rst.BOF And rst.EOF Then
        MsgBox "未找到刚创建的订单，请检查导入是否成功", vbExclamation
        rst.Close
        GoTo CleanUp
    End If
    
    lngOrderID = rst.Fields("OrderID").Value
    rst.Close
    
    ' ── 5. 执行批量生产 + 排版 ──
    ' 锁定屏幕加速
    App.Frame.ProjectBarUpdating = False
    App.DisableUndo = True
    
    ' 调用 CDM 加工引擎
    Call g_Make_Master(CStr(lngOrderID))
    
    ' ── 6. 完成 ──
    MsgBox "自动化生产排版完成！" & vbCrLf & vbCrLf & _
           "订单: " & sJobName & vbCrLf & _
           "后续可在 AlphaCAM 中查看并输出 NC", vbInformation, "自动化生产排版"
    GoTo CleanUp
    
EH:
    MsgBox "错误: " & Err.Description & vbCrLf & _
           "位置: " & Erl, vbCritical, "自动化生产排版"
    
CleanUp:
    Set rst = Nothing
    Set FSO = Nothing
    App.Frame.ProjectBarUpdating = True
    App.DisableUndo = False
End Sub
