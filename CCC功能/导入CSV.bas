Option Explicit

' ==============================================================================
' 导入CSV — 独立版（不依赖 CDM 工程）
' ==============================================================================
' 从 CSV 导入门板数据到 CDM 数据库，自动创建订单和明细
' 用法: 点击 CCC功能 → 导入CSV → 输入CSV文件名
' ==============================================================================

' CSV 列索引（1-based 列号）
Private Const COL_STYLE_NAME    As Integer = 0   ' 列1: 门板类型
Private Const COL_WIDTH         As Integer = 1   ' 列2: 宽
Private Const COL_HEIGHT        As Integer = 2   ' 列3: 高
Private Const COL_QUANTITY      As Integer = 3   ' 列4: 数量
Private Const COL_GROUP_ID      As Integer = 4   ' 列5: 组编号
Private Const COL_CUSTOMER      As Integer = 5   ' 列6: 客户名
Private Const COL_CUSTOM_1      As Integer = 7   ' 列8: 自定义1（开启方向）
Private Const COL_CUSTOM_2      As Integer = 8   ' 列9: 自定义2（终端地址）
Private Const COL_ORDER_REF     As Integer = 9   ' 列10: 订单号
Private Const COL_REMARK        As Integer = 11  ' 列12: 生产注释
Private Const COL_MATERIAL      As Integer = 12  ' 列13: 材料

Private Const DEF_MATERIAL_NAME As String = "开料机3000mm"
Private Const DEF_MATERIAL_THK  As Double = 18
Private Const DEF_MATERIAL_W    As Double = 1220
Private Const DEF_MATERIAL_L    As Double = 3000
Private Const DEF_DB_PATH       As String = "D:\2016\LICOMDAT\CDM Data\CDM.mdb"
Private Const DEF_DB_CONN       As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="


' ============================================================================
' 入口
' ============================================================================
Public Sub 导入CSV()
    '
    Dim sCSVPath As String
    Dim sJobName As String
    Dim sFolder  As String
    Dim sTemp    As String
    
    On Error GoTo EH
    
    sCSVPath = InputBox( _
        "请输入 CSV 文件名（含 .csv 扩展名）" & vbCrLf & _
        "目录: C:\Users\C\Desktop\2026优化表\", _
        "导入CSV", "7-10中林SPC婷兰灰.csv")
    
    If sCSVPath = "" Then Exit Sub
    
    sFolder = "C:\Users\C\Desktop\2026优化表\"
    If InStr(sCSVPath, "\") = 0 And InStr(sCSVPath, "/") = 0 Then
        sCSVPath = sFolder & sCSVPath
    End If
    
    If Dir(sCSVPath) = "" Then
        MsgBox "文件不存在:" & vbCrLf & sCSVPath, vbExclamation: Exit Sub
    End If
    
    sTemp = sCSVPath
    Do While InStr(sTemp, "\") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "\") + 1): Loop
    Do While InStr(sTemp, "/") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "/") + 1): Loop
    If LCase(Right$(sTemp, 4)) = ".csv" Then
        sJobName = Left$(sTemp, Len(sTemp) - 4)
    Else
        sJobName = sTemp
    End If
    
    Call ImportCSV sCSVPath, sJobName
    Exit Sub
    
EH:
    MsgBox "错误: " & Err.Description, vbCritical, "导入CSV"
End Sub

' ============================================================================
' 带参入口（供 modAutoProd 调用）
' ============================================================================
Public Sub 导入CSV文件(ByVal sCSVPath As String, ByVal sJobName As String)
    Call ImportCSV(sCSVPath, sJobName)
End Sub


' ============================================================================
' 核心导入
' ============================================================================
Private Sub ImportCSV(ByVal sCSVPath As String, ByVal sJobName As String)
    '
    Dim conn     As Object  ' ADODB.Connection
    Dim sConn    As String
    Dim lngOrderID As Long
    Dim lngRow   As Long
    Dim lngOK    As Long
    Dim sLine    As String
    Dim vFields  As Variant
    
    On Error GoTo EH
    
    ' ── 连接数据库 ──
    Set conn = CreateObject("ADODB.Connection")
    conn.Open DEF_DB_CONN & DEF_DB_PATH
    
    ' ── 确保客户存在 ──
    Call glng_EnsureCustomer(conn, "默认客户")
    
    ' ── 创建订单 ──
    lngOrderID = glng_CreateOrder(conn, sJobName, 1)
    If lngOrderID <= 0 Then MsgBox "创建订单失败", vbCritical: GoTo CleanUp
    
    ' ── 读取 CSV ──
    Dim iFile As Integer: iFile = FreeFile
    Open sCSVPath For Input As #iFile
    
    If Not EOF(iFile) Then Line Input #iFile, sLine  ' 跳过标题
    
    lngRow = 0: lngOK = 0
    
    Do While Not EOF(iFile)
        Line Input #iFile, sLine
        lngRow = lngRow + 1
        sLine = Trim$(sLine)
        If sLine = "" Then GoTo NextLine
        
        vFields = SplitCSVLine(sLine)
        If UBound(vFields) < 3 Then GoTo NextLine
        
        Dim sStyle As String, dblW As Double, dblH As Double, lngQ As Long
        Dim sMat As String, sCust As String, sRef As String, sRemark As String
        Dim sC1 As String, sC2 As String
        
        sStyle = Trim$(GetField(vFields, COL_STYLE_NAME, ""))
        dblW = Val(GetField(vFields, COL_WIDTH, "0"))
        dblH = Val(GetField(vFields, COL_HEIGHT, "0"))
        lngQ = Val(GetField(vFields, COL_QUANTITY, "1"))
        sMat = Trim$(GetField(vFields, COL_MATERIAL, ""))
        If sMat = "" Then sMat = DEF_MATERIAL_NAME
        sCust = Trim$(GetField(vFields, COL_CUSTOMER, ""))
        sRef = Trim$(GetField(vFields, COL_ORDER_REF, ""))
        sRemark = Trim$(GetField(vFields, COL_REMARK, ""))
        sC1 = Trim$(GetField(vFields, COL_CUSTOM_1, ""))
        sC2 = Trim$(GetField(vFields, COL_CUSTOM_2, ""))
        
        If dblW <= 0 Or dblH <= 0 Then GoTo NextLine
        If lngQ <= 0 Then lngQ = 1
        
        ' 确保门型和材料存在
        Call glng_EnsureStyle conn, sStyle
        Call glng_EnsureMaterial conn, sMat
        
        ' 插入明细
        Call InsertDetail conn, lngOrderID, sStyle, dblW, dblH, lngQ, _
                          sMat, sCust, sRef, sRemark, sC1, sC2
        lngOK = lngOK + 1
NextLine:
    Loop
    
    Close #iFile
    
    ' 保存 OrderID 到注册表（供自动排版读取）
    SaveSetting "CCC", "AutoProd", "LastOrderID", CStr(lngOrderID)
    
    MsgBox "导入完成！" & vbCrLf & _
           "订单: " & sJobName & " (ID: " & lngOrderID & ")" & vbCrLf & _
           "导入: " & lngOK & " 条", vbInformation, "导入CSV"
    GoTo CleanUp
    
EH:
    MsgBox "第 " & lngRow & " 行错误: " & Err.Description, vbExclamation, "导入CSV"
    
CleanUp:
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
    Close #iFile
End Sub


' ============================================================================
' 数据库辅助
' ============================================================================
Private Function glng_EnsureCustomer(conn As Object, ByVal sName As String) As Long
    Dim rst As Object: Set rst = CreateObject("ADODB.Recordset")
    rst.Open "SELECT CustomerID FROM AD_CUSTOMERS WHERE Name='" & gs_FixSQL(sName) & "'", _
             conn, 0, 1  ' adOpenForwardOnly, adLockReadOnly
    If Not (rst.BOF And rst.EOF) Then
        glng_EnsureCustomer = rst.Fields("CustomerID")
    Else
        rst.Close: conn.Execute "INSERT INTO AD_CUSTOMERS (Name) VALUES ('" & gs_FixSQL(sName) & "')"
        Set rst = conn.Execute("SELECT @@IDENTITY AS NewID")
        glng_EnsureCustomer = rst.Fields("NewID")
    End If
    rst.Close: Set rst = Nothing
End Function

Private Function glng_CreateOrder(conn As Object, ByVal sName As String, ByVal lngCustID As Long) As Long
    Dim rst As Object, lngRet As Long
    conn.Execute "INSERT INTO AD_ORDERS (JobName, CustomerID, OrderDate) VALUES ('" & _
                 gs_FixSQL(sName) & "', " & lngCustID & ", Date())", lngRet
    If lngRet > 0 Then Set rst = conn.Execute("SELECT @@IDENTITY AS NewID"): glng_CreateOrder = rst.Fields("NewID"): rst.Close
End Function

Private Sub glng_EnsureStyle(conn As Object, ByVal sName As String)
    Dim rst As Object, lngRet As Long
    If sName = "" Then Exit Sub
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open "SELECT PK FROM AD_DOOR_TYPES WHERE TypeID='" & gs_FixSQL(sName) & "'", conn, 0, 1
    If rst.BOF And rst.EOF Then rst.Close: conn.Execute "INSERT INTO AD_DOOR_TYPES (TypeID, StyleNumber) VALUES ('" & gs_FixSQL(sName) & "', 900)", lngRet
    rst.Close: Set rst = Nothing
End Sub

Private Sub glng_EnsureMaterial(conn As Object, ByVal sName As String)
    Dim rst As Object, lngRet As Long
    If sName = "" Then Exit Sub
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open "SELECT Name FROM AD_MATERIALS WHERE Name='" & gs_FixSQL(sName) & "'", conn, 0, 1
    If rst.BOF And rst.EOF Then rst.Close: conn.Execute "INSERT INTO AD_MATERIALS (Name, Thickness, SheetWidth, SheetLength) VALUES ('" & gs_FixSQL(sName) & "', " & DEF_MATERIAL_THK & ", " & DEF_MATERIAL_W & ", " & DEF_MATERIAL_L & ")", lngRet
    rst.Close: Set rst = Nothing
End Sub

Private Sub InsertDetail(conn As Object, ByVal lngOrderID As Long, _
    ByVal sStyle As String, ByVal dblW As Double, ByVal dblH As Double, ByVal lngQ As Long, _
    ByVal sMat As String, ByVal sCust As String, ByVal sRef As String, _
    ByVal sRemark As String, ByVal sC1 As String, ByVal sC2 As String)
    Dim lngRet As Long
    conn.Execute "INSERT INTO AD_ORDER_DETAILS " & _
        "(OrderID, TypeName, StyleNumber, Quantity, Width, Length, " & _
        "Material, ProductionComment, " & _
        "CSV_CustomerName, CSV_OrderNumber, CSV_ItemNumber, " & _
        "CustomField1, CustomField2) " & _
        "VALUES (" & lngOrderID & ",'" & gs_FixSQL(sStyle) & "',900," & _
        lngQ & "," & dblW & "," & dblH & ",'" & gs_FixSQL(sMat) & "'," & _
        "'" & gs_FixSQL(sRemark) & "','" & gs_FixSQL(sCust) & "'," & _
        "'" & gs_FixSQL(sRef) & "','" & gs_FixSQL(sRef) & "'," & _
        "'" & gs_FixSQL(sC1) & "','" & gs_FixSQL(sC2) & "')", lngRet
End Sub


' ============================================================================
' SQL 转义
' ============================================================================
Private Function gs_FixSQL(ByVal sText As String) As String
    If sText = "" Then gs_FixSQL = "": Exit Function
    gs_FixSQL = Replace(sText, "'", "''")
End Function


' ============================================================================
' CSV 解析
' ============================================================================
Private Function SplitCSVLine(ByVal sLine As String) As Variant
    Dim vResult() As String, iField As Long, iPos As Long, sField As String, bQuote As Boolean
    ReDim vResult(0 To 20)
    For iPos = 1 To Len(sLine)
        Dim sCh As String: sCh = Mid$(sLine, iPos, 1)
        If bQuote Then
            If sCh = """" Then
                If iPos < Len(sLine) And Mid$(sLine, iPos + 1, 1) = """" Then
                    sField = sField & """": iPos = iPos + 1
                Else: bQuote = False: End If
            Else: sField = sField & sCh: End If
        Else
            If sCh = """" Then: bQuote = True
            ElseIf sCh = "," Then: vResult(iField) = sField: iField = iField + 1: sField = ""
            Else: sField = sField & sCh: End If
        End If
    Next
    vResult(iField) = sField: ReDim Preserve vResult(0 To iField)
    SplitCSVLine = vResult
End Function

Private Function GetField(ByRef vFields As Variant, ByVal iIdx As Integer, ByVal sDef As String) As String
    If iIdx >= 0 And iIdx <= UBound(vFields) Then GetField = vFields(iIdx) Else GetField = sDef
End Function
