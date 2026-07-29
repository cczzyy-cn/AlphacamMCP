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
    
    ' ── 2. 导入 CSV → 创建订单（直接数据库操作）──
    Dim lngOrderID As Long
    App.Frame.ShowProgressBox "自动化生产排版", "正在导入 CSV ..."
    DoEvents
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\2016\LICOMDAT\CDM Data\CDM.mdb"
    
    ' 创建订单
    conn.Execute "INSERT INTO AD_ORDERS (JobName,CustomerID,OrderDate) VALUES ('" & _
                 FixSQL(sJobName) & "',1,Date())"
    Dim rs As Object
    Set rs = conn.Execute("SELECT @@IDENTITY AS NewID")
    lngOrderID = rs.Fields("NewID")
    rs.Close
    
    ' 读取 CSV 并插入明细
    Dim iFile As Integer: iFile = FreeFile
    Dim sLine As String, vF As Variant, lngRow As Long, lngOK As Long
    Open sCSVPath For Input As #iFile
    If Not EOF(iFile) Then Line Input #iFile, sLine  ' 跳过标题行
    
    Do While Not EOF(iFile)
        Line Input #iFile, sLine: lngRow = lngRow + 1
        sLine = Trim$(sLine): If sLine = "" Then GoTo NextLine
        vF = SplitCSVLine(sLine)
        If UBound(vF) < 3 Then GoTo NextLine
        
        Dim sTp As String, w As Double, h As Double, q As Long
        Dim sMat As String, sCu As String, sRf As String, sRm As String
        Dim sC1 As String, sC2 As String
        
        sTp = Trim$(GetF(vF, 0, "")): w = Val(GetF(vF, 1, "0"))
        h = Val(GetF(vF, 2, "0")): q = Val(GetF(vF, 3, "1"))
        sMat = Trim$(GetF(vF, 12, "")): If sMat = "" Then sMat = "开料机3000mm"
        sCu = Trim$(GetF(vF, 5, "")): sRf = Trim$(GetF(vF, 9, ""))
        sRm = Trim$(GetF(vF, 11, "")): sC1 = Trim$(GetF(vF, 7, ""))
        sC2 = Trim$(GetF(vF, 8, ""))
        If w <= 0 Or h <= 0 Then GoTo NextLine
        If q <= 0 Then q = 1
        
        ' 确保门型 + 材料存在（StyleNumber=900，StyleName=TypeName）
        conn.Execute "INSERT INTO AD_DOOR_TYPES (TypeID,StyleNumber) SELECT '" & _
                     FixSQL(sTp) & "',900 WHERE NOT EXISTS(SELECT 1 FROM AD_DOOR_TYPES WHERE TypeID='" & FixSQL(sTp) & "')"
        conn.Execute "INSERT INTO AD_MATERIALS (Name,Thickness,SheetWidth,SheetLength) SELECT '" & _
                     FixSQL(sMat) & "',18,1220,3000 WHERE NOT EXISTS(SELECT 1 FROM AD_MATERIALS WHERE Name='" & FixSQL(sMat) & "')"
        
        ' 插入明细
        conn.Execute "INSERT INTO AD_ORDER_DETAILS " & _
            "(OrderID,TypeName,StyleName,StyleNumber,Quantity,Width,Length," & _
            "Material,ProductionComment," & _
            "CSV_CustomerName,CSV_OrderNumber,CSV_ItemNumber," & _
            "CustomField1,CustomField2) VALUES (" & _
            lngOrderID & ",'" & FixSQL(sTp) & "','" & FixSQL(sTp) & "',900," & _
            q & "," & w & "," & h & ",'" & FixSQL(sMat) & "'," & _
            "'" & FixSQL(sRm) & "','" & FixSQL(sCu) & "'," & _
            "'" & FixSQL(sRf) & "','" & FixSQL(sRf) & "'," & _
            "'" & FixSQL(sC1) & "','" & FixSQL(sC2) & "')"
        lngOK = lngOK + 1
NextLine:
    Loop
    Close #iFile
    conn.Close: Set conn = Nothing
    
    App.Frame.SetProgressText "导入完成: " & lngOK & " 条记录"
    
    ' ── 3. 调用 CDM 加工引擎 ──
    App.Frame.CloseProgressBox
    MsgBox "CSV 导入完成！订单: " & sJobName & " (ID: " & lngOrderID & ")" & vbCrLf & vbCrLf & _
           "请通过 CDM 菜单 → Processing → Orders → 选择该订单 → Load to Production Queue" & vbCrLf & _
           "→ Run Production 完成批量生产+排版。", vbInformation, "自动化生产排版"
    Exit Sub
    ' (注: g_Make_Master 在 CCC 工程中无法直接调用，需通过 CDM UI 操作)
    
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

' ============================================================================
' 工具函数
' ============================================================================
Private Function FixSQL(ByVal s As String) As String
    If s = "" Then FixSQL = "": Else FixSQL = Replace(s, "'", "''")
End Function

Private Function SplitCSVLine(ByVal sLine As String) As Variant
    Dim v() As String, idx As Long, i As Long, f As String, bQ As Boolean
    ReDim v(0 To 20)
    For i = 1 To Len(sLine)
        Dim c As String: c = Mid$(sLine, i, 1)
        If bQ Then
            If c = """" Then
                If i < Len(sLine) And Mid$(sLine, i + 1, 1) = """" Then
                    f = f & """": i = i + 1
                Else: bQ = False: End If
            Else: f = f & c: End If
        Else
            If c = """" Then: bQ = True
            ElseIf c = "," Then: v(idx) = f: idx = idx + 1: f = ""
            Else: f = f & c: End If
        End If
    Next
    v(idx) = f: ReDim Preserve v(0 To idx): SplitCSVLine = v
End Function

Private Function GetF(ByRef v As Variant, ByVal i As Integer, ByVal d As String) As String
    If i >= 0 And i <= UBound(v) Then GetF = v(i) Else GetF = d
End Function
