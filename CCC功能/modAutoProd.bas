Option Explicit

' API 文件选择对话框
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_HIDEREADONLY  As Long = &H4

' ==============================================================================
' 自动化生产排版 — 弹窗选CSV → 导入 → 批量生产 → 排版
' ==============================================================================
' 用法: 点击 CCC功能 → 自动化生产排版 → 选择CSV文件
' ==============================================================================

Private Const MAT_THK  As Double = 18
Private Const MAT_W    As Double = 1220
Private Const MAT_L    As Double = 3000
Private Const MAT_NAME As String = "开料机3000mm"
Private Const DB_PATH  As String = "D:\2016\LICOMDAT\CDM Data\CDM.mdb"
Private Const DB_CONN  As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

' CSV 列索引（0-based）
Private Const C_STYLE  As Integer = 0   ' 列1
Private Const C_W      As Integer = 1   ' 列2
Private Const C_H      As Integer = 2   ' 列3
Private Const C_QTY    As Integer = 3   ' 列4
Private Const C_CUST   As Integer = 5   ' 列6
Private Const C_REF    As Integer = 9   ' 列10
Private Const C_REMARK As Integer = 11  ' 列12
Private Const C_MAT    As Integer = 12  ' 列13
Private Const C_C1     As Integer = 7   ' 列8
Private Const C_C2     As Integer = 8   ' 列9


' ============================================================================
' 主入口
' ============================================================================
Public Sub 自动化生产排版()
    '
    Dim sCSVPath As String
    Dim sJobName As String
    Dim sTemp    As String
    
    On Error GoTo EH
    
    ' ── 1. 选择 CSV 文件 ──
    sCSVPath = ShowOpenFileDialog("C:\Users\C\Desktop\2026优化表", "CSV 文件|*.csv|所有文件|*.*")
    If sCSVPath = "" Then Exit Sub
    
    ' 取文件名（不含 .csv）作为订单名
    sTemp = sCSVPath
    Do While InStr(sTemp, "\") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "\") + 1): Loop
    Do While InStr(sTemp, "/") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "/") + 1): Loop
    If LCase(Right$(sTemp, 4)) = ".csv" Then
        sJobName = Left$(sTemp, Len(sTemp) - 4)
    Else
        sJobName = sTemp
    End If
    
    ' ── 2. 导入 CSV → 创建订单 ──
    Dim lngOrderID As Long
    App.Frame.ShowProgressBox "自动化生产排版", "正在导入 CSV ..."
    DoEvents
    
    ' 通过 CDM 工程的 BatchImport 导入（确保数据库连接稳定）
    On Error Resume Next
    Application.Run "CDM.BatchImport.Run", sCSVPath, sJobName
    If Err.Number <> 0 Then
        Dim sErr2 As String: sErr2 = Err.Description
        On Error GoTo EH
        App.Frame.CloseProgressBox
        If InStr(sErr2, "Object required") Or InStr(sErr2, "未找到") Then
            MsgBox "BatchImport 模块未安装！" & vbCrLf & vbCrLf & _
                   "请先在 CDM 工程的 VBA 编辑器中导入:" & vbCrLf & _
                   "CDM功能/BatchImport.bas", vbCritical, "自动化生产排版"
        Else
            MsgBox "CSV 导入失败:" & vbCrLf & sErr2, vbCritical, "自动化生产排版"
        End If
        Exit Sub
    End If
    On Error GoTo EH
    
    ' 从注册表读取 OrderID（BatchImport 已保存）
    lngOrderID = CLng(GetSetting("CCC", "AutoProd", "LastOrderID", "0"))
    
    ' ── 3. 调用 CDM 加工引擎 ──
    App.Frame.ShowProgressBox "自动化生产排版", "正在执行批量生产 + 排版 ..."
    DoEvents
    
    Application.Run "CDM.g_Make_Master", CStr(lngOrderID)
    
    App.Frame.CloseProgressBox
    
    ' ── 4. 完成 ──
    Dim bOK As Boolean
    bOK = CBool(GetSetting("LICOM AlphaDOOR", "Nest Parameters", "Nest Completed", 0))
    
    If bOK Then
        MsgBox "自动化生产排版完成！" & vbCrLf & vbCrLf & _
               "订单: " & sJobName & vbCrLf & _
               "可在 AlphaCAM 中查看结果并输出 NC", vbInformation
    Else
        MsgBox "生产排版可能未完全成功，请检查 AlphaCAM 结果。" & vbCrLf & vbCrLf & _
               "常见原因:" & vbCrLf & _
               "  - 门型刀路未配置 (AD_DOOR_PATHS)" & vbCrLf & _
               "  - 材料参数不正确" & vbCrLf & _
               "  - 后处理器未选择", vbExclamation
    End If
    Exit Sub
    
EH:
    App.Frame.CloseProgressBox
    MsgBox "错误: " & Err.Description, vbCritical, "自动化生产排版"
End Sub


' ============================================================================
' 文件选择对话框
' ============================================================================
Private Function ShowOpenFileDialog(ByVal sInitialDir As String, _
                                    ByVal sFilter As String) As String
    '
    Dim ofn As OPENFILENAME
    Dim sFile As String
    Dim lRet As Long
    
    ' 初始化结构体
    sFile = String$(260, vbNullChar)
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = App.Frame.WindowHandle
    ofn.lpstrFilter = sFilter
    ofn.lpstrFile = sFile
    ofn.nMaxFile = Len(sFile)
    ofn.lpstrInitialDir = sInitialDir
    ofn.lpstrTitle = "选择 CSV 文件"
    ofn.flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    ofn.lpstrDefExt = "csv"
    
    lRet = GetOpenFileName(ofn)
    
    If lRet Then
        ShowOpenFileDialog = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
    End If
End Function
