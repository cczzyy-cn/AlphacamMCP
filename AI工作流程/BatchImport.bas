Attribute VB_Name = "BatchImport"
Option Explicit

' ============================================================================
' BatchImport — CDM 批量导入模块
' ============================================================================
' 功能: 从 CSV 文件批量导入门板数据到 CDM 数据库，自动创建订单和明细
' 依赖: CDM.arb 已加载（提供 gdb_CDM、gbln_ConnectToDB、gs_FixSQL）
' 用法: CDM.BatchImport.Run "C:\path\to\file.csv", "订单名称"
' ============================================================================

' CSV 列索引（匹配 CDM 导入配置：1-based 列号）
' 列1=门板类型  列2=宽  列3=高  列4=数量  列5=组编号
' 列6=客户名    列7=   列8=自定义1  列9=自定义2  列10=订单号
' 列11=        列12=生产注释  列13=材料  列14~19=贴面纹理/后处理/旋转等
Private Const COL_STYLE_NAME    As Integer = 0   ' 列1: 门板类型（造型名称）
Private Const COL_WIDTH         As Integer = 1   ' 列2: 宽
Private Const COL_HEIGHT        As Integer = 2   ' 列3: 高
Private Const COL_QUANTITY      As Integer = 3   ' 列4: 数量
Private Const COL_GROUP_ID      As Integer = 4   ' 列5: 组编号（原颜色）
Private Const COL_CUSTOMER      As Integer = 5   ' 列6: 客户名称
Private Const COL_CUSTOM_1      As Integer = 7   ' 列8: 自定义字符串1（开启方向）
Private Const COL_CUSTOM_2      As Integer = 8   ' 列9: 自定义字符串2（终端地址）
Private Const COL_ORDER_REF     As Integer = 9   ' 列10: 订单号（板件码）
Private Const COL_REMARK        As Integer = 11  ' 列12: 生产注释（备注）
Private Const COL_MATERIAL      As Integer = 12  ' 列13: 材料（M列，空时用默认材料）

' 默认材料
Private Const DEF_MATERIAL_NAME As String = "开料机3000mm"
Private Const DEF_MATERIAL_THK  As Double = 18
Private Const DEF_MATERIAL_W    As Double = 1220
Private Const DEF_MATERIAL_L    As Double = 3000


' ============================================================================
' 主入口 — 从指定 CSV 文件导入
' ============================================================================
Public Sub Run(Optional sCSVPath As String = "", _
               Optional sJobName As String = "")
    '
    Dim FSO As New Scripting.FileSystemObject
    Dim sDefaultCSV As String
    Dim sDefaultJob As String
    
    ' 默认值：从文件名推断
    If sCSVPath = "" Then
        sDefaultCSV = "C:\Users\C\Desktop\2026优化表\7-10中林SPC婷兰灰.csv"
        sCSVPath = sDefaultCSV
    End If
    
    If sJobName = "" Then
        ' 从 CSV 文件名去掉 .csv 作为订单名
        sDefaultJob = FSO.GetBaseName(sCSVPath)
        sJobName = sDefaultJob
    End If
    
    ' 检查文件是否存在
    If Not FSO.FileExists(sCSVPath) Then
        MsgBox "CSV 文件不存在: " & sCSVPath, vbExclamation, "BatchImport"
        Exit Sub
    End If
    
    ' 执行导入
    Call ImportCSV sCSVPath, sJobName
    
    Set FSO = Nothing
End Sub


' ============================================================================
' 核心：导入单个 CSV 到 CDM 数据库
' ============================================================================
Private Sub ImportCSV(ByVal sCSVPath As String, ByVal sJobName As String)
    '
    Dim FSO As New Scripting.FileSystemObject
    Dim ts As TextStream
    Dim sLine As String
    Dim vFields As Variant
    Dim lngOrderID As Long
    Dim lngCustomerID As Long
    Dim lngRow As Long
    Dim lngImported As Long
    Dim lngSkipped As Long
    Dim sErrMsg As String
    
    On Error GoTo EH
    
    ' ── 连接数据库 ──
    If Not gbln_ConnectToDB() Then
        MsgBox "无法连接到 CDM 数据库，请检查 CDM.udl 配置。", vbCritical, "BatchImport"
        GoTo CleanUp
    End If
    
    ' ── 获取或创建客户 ──
    lngCustomerID = glng_GetOrCreateCustomer("默认客户")
    
    ' ── 创建订单 ──
    lngOrderID = glng_CreateOrder(sJobName, lngCustomerID)
    If lngOrderID <= 0 Then
        MsgBox "创建订单失败！", vbCritical, "BatchImport"
        GoTo CleanUp
    End If
    
    ' ── 读取 CSV 并导入 ──
    Set ts = FSO.OpenTextFile(sCSVPath, ForReading, False)
    
    ' 跳过标题行
    If Not ts.AtEndOfStream Then ts.SkipLine
    
    lngRow = 0
    lngImported = 0
    lngSkipped = 0
    
    Do While Not ts.AtEndOfStream
        sLine = Trim$(ts.ReadLine)
        lngRow = lngRow + 1
        
        ' 跳过空行
        If sLine = "" Then GoTo NextRow
        
        ' 解析 CSV 行
        vFields = SplitCSVLine(sLine)
        
        ' 验证必要字段
        If UBound(vFields) < 3 Then
            lngSkipped = lngSkipped + 1
            GoTo NextRow
        End If
        
        ' 提取字段
        Dim sStyleName    As String
        Dim dblWidth      As Double
        Dim dblHeight     As Double
        Dim lngQty        As Long
        Dim sMaterial     As String
        Dim sCustomer     As String
        Dim sOrderRef     As String
        Dim sPartCode     As String
        Dim sRemark       As String
        Dim sCustom1      As String  ' 自定义字符串1（开启方向）
        Dim sCustom2      As String  ' 自定义字符串2（终端地址）
        
        sStyleName = Trim$(GetField(vFields, COL_STYLE_NAME, ""))
        dblWidth = Val(GetField(vFields, COL_WIDTH, "0"))
        dblHeight = Val(GetField(vFields, COL_HEIGHT, "0"))
        lngQty = Val(GetField(vFields, COL_QUANTITY, "1"))
        
        ' 材料从第13列(M列)读取，空时用默认材料
        sMaterial = Trim$(GetField(vFields, COL_MATERIAL, ""))
        If sMaterial = "" Then sMaterial = DEF_MATERIAL_NAME
        
        sCustomer = Trim$(GetField(vFields, COL_CUSTOMER, ""))
        sOrderRef = Trim$(GetField(vFields, COL_ORDER_REF, ""))
        sPartCode = Trim$(GetField(vFields, COL_ORDER_REF, ""))  ' 板件码=订单号列
        sRemark = Trim$(GetField(vFields, COL_REMARK, ""))
        sCustom1 = Trim$(GetField(vFields, COL_CUSTOM_1, ""))    ' 列8: 开启方向
        sCustom2 = Trim$(GetField(vFields, COL_CUSTOM_2, ""))    ' 列9: 终端地址
        
        ' 跳过无效行
        If dblWidth <= 0 Or dblHeight <= 0 Then
            lngSkipped = lngSkipped + 1
            GoTo NextRow
        End If
        If lngQty <= 0 Then lngQty = 1
        
        ' 插入到 AD_ORDER_DETAILS
        Call InsertOrderDetail lngOrderID, sStyleName, dblWidth, dblHeight, _
                               lngQty, sMaterial, sCustomer, sOrderRef, sPartCode, sRemark, _
                               sCustom1, sCustom2
        lngImported = lngImported + 1
        
NextRow:
    Loop
    
    ts.Close
    
    ' ── 完成 ──
    sErrMsg = "导入完成！" & vbCrLf & _
              "  订单: " & sJobName & " (ID: " & lngOrderID & ")" & vbCrLf & _
              "  材料: " & sMaterial & vbCrLf & _
              "  成功导入: " & lngImported & " 条" & vbCrLf & _
              "  跳过: " & lngSkipped & " 行"
    MsgBox sErrMsg, vbInformation, "BatchImport"
    GoTo CleanUp
    
EH:
    sErrMsg = "在第 " & lngRow & " 行发生错误:" & vbCrLf & _
              Err.Description & vbCrLf & vbCrLf & _
              "已成功导入 " & lngImported & " 条记录。"
    MsgBox sErrMsg, vbExclamation, "BatchImport"
    
CleanUp:
    Set ts = Nothing
    Set FSO = Nothing
End Sub


' ============================================================================
' 获取或创建客户
' ============================================================================
Private Function glng_GetOrCreateCustomer(ByVal sName As String) As Long
    '
    Dim rst As ADODB.Recordset
    Dim lngRet As Long
    
    ' 查找现有客户
    Set rst = New ADODB.Recordset
    rst.Open "SELECT CustomerID FROM AD_CUSTOMERS WHERE Name='" & gs_FixSQL(sName) & "'", _
             gdb_CDM, adOpenForwardOnly, adLockReadOnly
    
    If Not (rst.BOF And rst.EOF) Then
        glng_GetOrCreateCustomer = rst.Fields("CustomerID").Value
    Else
        rst.Close
        ' 创建新客户
        gdb_CDM.Execute "INSERT INTO AD_CUSTOMERS (Name) VALUES ('" & gs_FixSQL(sName) & "')", lngRet
        Set rst = gdb_CDM.Execute("SELECT @@IDENTITY AS NewID")
        glng_GetOrCreateCustomer = rst.Fields("NewID").Value
    End If
    
    rst.Close
    Set rst = Nothing
End Function


' ============================================================================
' 创建订单
' ============================================================================
Private Function glng_CreateOrder(ByVal sJobName As String, _
                                  ByVal lngCustomerID As Long) As Long
    '
    Dim rst As ADODB.Recordset
    Dim lngRet As Long
    Dim sSQL As String
    
    sSQL = "INSERT INTO AD_ORDERS (JobName, CustomerID, OrderDate) VALUES ('" & _
           gs_FixSQL(sJobName) & "', " & lngCustomerID & ", Date())"
    
    gdb_CDM.Execute sSQL, lngRet
    
    If lngRet > 0 Then
        Set rst = gdb_CDM.Execute("SELECT @@IDENTITY AS NewID")
        If Not (rst.BOF And rst.EOF) Then
            glng_CreateOrder = rst.Fields("NewID").Value
        End If
        rst.Close
    End If
    
    Set rst = Nothing
End Function


' ============================================================================
' 插入订单明细
' ============================================================================
Private Sub InsertOrderDetail(ByVal lngOrderID As Long, _
                              ByVal sStyleName As String, _
                              ByVal dblWidth As Double, _
                              ByVal dblHeight As Double, _
                              ByVal lngQty As Long, _
                              ByVal sMaterial As String, _
                              ByVal sCustomer As String, _
                              ByVal sOrderRef As String, _
                              ByVal sPartCode As String, _
                              ByVal sRemark As String, _
                              ByVal sCustom1 As String, _
                              ByVal sCustom2 As String)
    '
    Dim lngRet As Long
    Dim sSQL As String
    
    ' 注册门型（StyleNumber=900 匹配 Make.mbln_ProcessPart）
    Call glng_EnsureStyle(sStyleName)
    
    ' 注册材料
    Call glng_EnsureMaterial(sMaterial)
    
    ' 插入明细（含自定义字段）
    sSQL = "INSERT INTO AD_ORDER_DETAILS " & _
           "(OrderID, TypeName, StyleNumber, Quantity, Width, Length, " & _
           "Material, ProductionComment, " & _
           "CSV_CustomerName, CSV_OrderNumber, CSV_ItemNumber, " & _
           "CustomField1, CustomField2) " & _
           "VALUES (" & lngOrderID & ", " & _
           "'" & gs_FixSQL(sStyleName) & "', 900, " & _
           "" & lngQty & ", " & _
           "" & dblWidth & ", " & _
           "" & dblHeight & ", " & _
           "'" & gs_FixSQL(sMaterial) & "', " & _
           "'" & gs_FixSQL(sRemark) & "', " & _
           "'" & gs_FixSQL(sCustomer) & "', " & _
           "'" & gs_FixSQL(sOrderRef) & "', " & _
           "'" & gs_FixSQL(sPartCode) & "', " & _
           "'" & gs_FixSQL(sCustom1) & "', " & _
           "'" & gs_FixSQL(sCustom2) & "')"
    
    gdb_CDM.Execute sSQL, lngRet
End Sub


' ============================================================================
' 确保门型存在（StyleNumber=900）
' ============================================================================
Private Sub glng_EnsureStyle(ByVal sTypeName As String)
    '
    Dim rst As ADODB.Recordset
    Dim lngRet As Long
    
    If sTypeName = "" Then Exit Sub
    
    Set rst = New ADODB.Recordset
    rst.Open "SELECT PK FROM AD_DOOR_TYPES WHERE TypeID='" & gs_FixSQL(sTypeName) & "'", _
             gdb_CDM, adOpenForwardOnly, adLockReadOnly
    
    If rst.BOF And rst.EOF Then
        rst.Close
        gdb_CDM.Execute "INSERT INTO AD_DOOR_TYPES (TypeID, StyleNumber) VALUES ('" & _
                        gs_FixSQL(sTypeName) & "', 900)", lngRet
    End If
    
    rst.Close
    Set rst = Nothing
End Sub


' ============================================================================
' 确保材料存在（默认: 开料机3000mm）
' ============================================================================
Private Sub glng_EnsureMaterial(ByVal sName As String)
    '
    Dim rst As ADODB.Recordset
    Dim lngRet As Long
    
    If sName = "" Then Exit Sub
    
    Set rst = New ADODB.Recordset
    rst.Open "SELECT Name FROM AD_MATERIALS WHERE Name='" & gs_FixSQL(sName) & "'", _
             gdb_CDM, adOpenForwardOnly, adLockReadOnly
    
    If rst.BOF And rst.EOF Then
        rst.Close
        gdb_CDM.Execute "INSERT INTO AD_MATERIALS (Name, Thickness, SheetWidth, SheetLength) " & _
                        "VALUES ('" & gs_FixSQL(sName) & "', " & _
                        DEF_MATERIAL_THK & ", " & _
                        DEF_MATERIAL_W & ", " & _
                        DEF_MATERIAL_L & ")", lngRet
    End If
    
    rst.Close
    Set rst = Nothing
End Sub


' ============================================================================
' CSV 行解析（支持引号包裹的字段）
' ============================================================================
Private Function SplitCSVLine(ByVal sLine As String) As Variant
    '
    Dim vResult() As String
    Dim i As Long
    Dim iPos As Long
    Dim iLen As Long
    Dim sField As String
    Dim bInQuote As Boolean
    Dim iField As Long
    
    ReDim vResult(0 To 20)
    iField = 0
    iLen = Len(sLine)
    sField = ""
    bInQuote = False
    
    For iPos = 1 To iLen
        Dim sCh As String
        sCh = Mid$(sLine, iPos, 1)
        
        If bInQuote Then
            If sCh = """" Then
                If iPos < iLen And Mid$(sLine, iPos + 1, 1) = """" Then
                    sField = sField & """"
                    iPos = iPos + 1
                Else
                    bInQuote = False
                End If
            Else
                sField = sField & sCh
            End If
        Else
            If sCh = """" Then
                bInQuote = True
            ElseIf sCh = "," Then
                If iField > UBound(vResult) Then
                    ReDim Preserve vResult(0 To iField + 10)
                End If
                vResult(iField) = sField
                iField = iField + 1
                sField = ""
            Else
                sField = sField & sCh
            End If
        End If
    Next
    
    ' 最后一个字段
    If iField > UBound(vResult) Then
        ReDim Preserve vResult(0 To iField + 10)
    End If
    vResult(iField) = sField
    
    ' 裁剪数组
    ReDim Preserve vResult(0 To iField)
    
    SplitCSVLine = vResult
End Function


' ============================================================================
' 安全获取数组字段
' ============================================================================
Private Function GetField(ByRef vFields As Variant, _
                          ByVal iIndex As Integer, _
                          ByVal sDefault As String) As String
    '
    If iIndex >= 0 And iIndex <= UBound(vFields) Then
        GetField = vFields(iIndex)
    Else
        GetField = sDefault
    End If
End Function
