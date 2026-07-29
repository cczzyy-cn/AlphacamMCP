Option Explicit

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
    Dim sFolder  As String
    Dim sTemp    As String
    
    On Error GoTo EH
    
    ' ── 1. 选择 CSV 文件 ──
    sCSVPath = InputBox( _
        "请输入 CSV 文件名（含 .csv 扩展名）" & vbCrLf & vbCrLf & _
        "目录: C:\Users\C\Desktop\2026优化表\", _
        "自动化生产排版", "7-10中林SPC婷兰灰.csv")
    
    If sCSVPath = "" Then Exit Sub
    
    sFolder = "C:\Users\C\Desktop\2026优化表\"
    If InStr(sCSVPath, "\") = 0 And InStr(sCSVPath, "/") = 0 Then
        sCSVPath = sFolder & sCSVPath
    End If
    
    If Dir(sCSVPath) = "" Then
        MsgBox "文件不存在!" & vbCrLf & sCSVPath, vbExclamation
        Exit Sub
    End If
    
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
    lngOrderID = ImportCSVAndCreateOrder(sCSVPath, sJobName)
    
    If lngOrderID <= 0 Then
        MsgBox "导入失败，无法创建订单", vbCritical
        Exit Sub
    End If
    
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
' 导入 CSV 并创建订单，返回 OrderID
' ============================================================================
Private Function ImportCSVAndCreateOrder(ByVal sCSVPath As String, _
                                         ByVal sJobName As String) As Long
    '
    Dim conn As Object      ' ADODB.Connection
    Dim lngOrderID As Long
    Dim lngRow As Long
    Dim lngOK As Long
    Dim sLine As String
    Dim vFields As Variant
    Dim iFile As Integer
    
    On Error GoTo EH
    
    ' 连接数据库
    Set conn = CreateObject("ADODB.Connection")
    conn.Open DB_CONN & DB_PATH
    
    ' 确保客户
    EnsureCustomer conn, "默认客户"
    
    ' 创建订单
    lngOrderID = CreateOrder(conn, sJobName, 1)
    If lngOrderID <= 0 Then GoTo CleanUp
    
    ' 读取 CSV
    iFile = FreeFile
    Open sCSVPath For Input As #iFile
    If Not EOF(iFile) Then Line Input #iFile, sLine  ' 跳过标题行
    
    lngRow = 0: lngOK = 0
    
    Do While Not EOF(iFile)
        Line Input #iFile, sLine
        lngRow = lngRow + 1
        sLine = Trim$(sLine)
        If sLine = "" Then GoTo NextLine
        
        vFields = SplitCSVLine(sLine)
        If UBound(vFields) < 3 Then GoTo NextLine
        
        Dim sStyle As String, dblW As Double, dblH As Double, lngQ As Long
        Dim sMat As String, sCust As String, sRef As String, sRem As String
        Dim sC1 As String, sC2 As String
        
        sStyle = Trim$(GetF(vFields, C_STYLE, ""))
        dblW = Val(GetF(vFields, C_W, "0"))
        dblH = Val(GetF(vFields, C_H, "0"))
        lngQ = Val(GetF(vFields, C_QTY, "1"))
        sMat = Trim$(GetF(vFields, C_MAT, ""))
        If sMat = "" Then sMat = MAT_NAME
        sCust = Trim$(GetF(vFields, C_CUST, ""))
        sRef = Trim$(GetF(vFields, C_REF, ""))
        sRem = Trim$(GetF(vFields, C_REMARK, ""))
        sC1 = Trim$(GetF(vFields, C_C1, ""))
        sC2 = Trim$(GetF(vFields, C_C2, ""))
        
        If dblW <= 0 Or dblH <= 0 Then GoTo NextLine
        If lngQ <= 0 Then lngQ = 1
        
        EnsureStyle conn, sStyle
        EnsureMaterial conn, sMat
        InsertDetail conn, lngOrderID, sStyle, dblW, dblH, lngQ, _
                     sMat, sCust, sRef, sRem, sC1, sC2
        lngOK = lngOK + 1
NextLine:
    Loop
    
    Close #iFile
    
    ImportCSVAndCreateOrder = lngOrderID
    GoTo CleanUp
    
EH:
    ImportCSVAndCreateOrder = 0
    MsgBox "第 " & lngRow & " 行错误: " & Err.Description, vbExclamation
    
CleanUp:
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    Set conn = Nothing
    Close #iFile
End Function


' ============================================================================
' 数据库辅助
' ============================================================================
Private Function CreateOrder(conn As Object, ByVal sName As String, _
                             ByVal lngCustID As Long) As Long
    Dim rst As Object, lngRet As Long
    conn.Execute "INSERT INTO AD_ORDERS (JobName,CustomerID,OrderDate) VALUES ('" & _
                 FixSQL(sName) & "'," & lngCustID & ",Date())", lngRet
    If lngRet > 0 Then
        Set rst = conn.Execute("SELECT @@IDENTITY AS NewID")
        CreateOrder = rst.Fields("NewID")
        rst.Close
    End If
End Function

Private Sub EnsureCustomer(conn As Object, ByVal sName As String)
    Dim rst As Object
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open "SELECT CustomerID FROM AD_CUSTOMERS WHERE Name='" & FixSQL(sName) & "'", _
             conn, 0, 1
    If rst.BOF And rst.EOF Then
        rst.Close: conn.Execute "INSERT INTO AD_CUSTOMERS (Name) VALUES ('" & FixSQL(sName) & "')"
    End If
    rst.Close
End Sub

Private Sub EnsureStyle(conn As Object, ByVal sName As String)
    Dim rst As Object, lngRet As Long
    If sName = "" Then Exit Sub
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open "SELECT PK FROM AD_DOOR_TYPES WHERE TypeID='" & FixSQL(sName) & "'", conn, 0, 1
    If rst.BOF And rst.EOF Then
        rst.Close
        conn.Execute "INSERT INTO AD_DOOR_TYPES (TypeID,StyleNumber) VALUES ('" & FixSQL(sName) & "',900)", lngRet
    End If
    rst.Close
End Sub

Private Sub EnsureMaterial(conn As Object, ByVal sName As String)
    Dim rst As Object, lngRet As Long
    If sName = "" Then Exit Sub
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open "SELECT Name FROM AD_MATERIALS WHERE Name='" & FixSQL(sName) & "'", conn, 0, 1
    If rst.BOF And rst.EOF Then
        rst.Close
        conn.Execute "INSERT INTO AD_MATERIALS (Name,Thickness,SheetWidth,SheetLength) VALUES ('" & _
                     FixSQL(sName) & "'," & MAT_THK & "," & MAT_W & "," & MAT_L & ")", lngRet
    End If
    rst.Close
End Sub

Private Sub InsertDetail(conn As Object, ByVal lngOrderID As Long, _
    ByVal sStyle As String, ByVal dblW As Double, ByVal dblH As Double, _
    ByVal lngQ As Long, ByVal sMat As String, ByVal sCust As String, _
    ByVal sRef As String, ByVal sRem As String, ByVal sC1 As String, ByVal sC2 As String)
    Dim lngRet As Long
    conn.Execute "INSERT INTO AD_ORDER_DETAILS " & _
        "(OrderID,TypeName,StyleNumber,Quantity,Width,Length," & _
        "Material,ProductionComment," & _
        "CSV_CustomerName,CSV_OrderNumber,CSV_ItemNumber," & _
        "CustomField1,CustomField2) VALUES (" & _
        lngOrderID & ",'" & FixSQL(sStyle) & "',900," & _
        lngQ & "," & dblW & "," & dblH & ",'" & FixSQL(sMat) & "'," & _
        "'" & FixSQL(sRem) & "','" & FixSQL(sCust) & "'," & _
        "'" & FixSQL(sRef) & "','" & FixSQL(sRef) & "'," & _
        "'" & FixSQL(sC1) & "','" & FixSQL(sC2) & "')", lngRet
End Sub

Private Function FixSQL(ByVal s As String) As String
    If s = "" Then FixSQL = "": Else FixSQL = Replace(s, "'", "''")
End Function


' ============================================================================
' CSV 解析
' ============================================================================
Private Function SplitCSVLine(ByVal sLine As String) As Variant
    Dim v() As String
    Dim iF As Long
    Dim i As Long
    Dim f As String
    Dim bQ As Boolean
    ReDim v(0 To 20)
    For i = 1 To Len(sLine)
        Dim c As String: c = Mid$(sLine, i, 1)
        If bQ Then
            If c = """" Then
                If i < Len(sLine) And Mid$(sLine, i + 1, 1) = """" Then
                    f = f & """"
                    i = i + 1
                Else
                    bQ = False
                End If
            Else
                f = f & c
            End If
        Else
            If c = """" Then
                bQ = True
            ElseIf c = "," Then
                v(iF) = f
                iF = iF + 1
                f = ""
            Else
                f = f & c
            End If
        End If
    Next
    v(iF) = f
    ReDim Preserve v(0 To iF)
    SplitCSVLine = v
End Function

Private Function GetF(ByRef v As Variant, ByVal i As Integer, ByVal d As String) As String
    If i >= 0 And i <= UBound(v) Then GetF = v(i) Else GetF = d
End Function
