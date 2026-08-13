' MCP-INSTALL-TEST: 2026-07-30 测试注释
Option Explicit

' ==============================================================================
' 自动化生产排版 — 弹窗选CSV → 导入 → 批量生产 → 排版（CDM 工程内）
' ==============================================================================
' 安装: 在 CDM 工程的 VBA 编辑器中导入此文件
' 用法: AlphaCAM 菜单 → CCC功能 → 自动化生产排版
' 说明: 由 CCC 功能菜单触发，代码在 CDM 工程中执行以直接调用 g_Make_Master
' ==============================================================================

' ============================================================================
' 主入口（CCC 功能菜单触发）
' ============================================================================
Public Sub AutoImportNest()
    ' 菜单入口：弹出 frmAutoNest 窗体（选 CSV → 导入 → 排版）
    On Error GoTo EH
    frmAutoNest.Show
    Exit Sub
EH:
    MsgBox "错误: " & Err.Description, vbCritical
End Sub


' ============================================================================
' 导入 CSV → 创建订单，返回 OrderID
' ============================================================================
Public Sub AutoImportNestWithParams(ByVal sCSVPath As String, _
                                  ByVal sCustomerName As String, _
                                  ByVal sMaterialName As String, _
                                  ByVal bRunNest As Boolean)
    ' 由 frmAutoNest 窗体调用的带参入口
    Dim sJobName As String, sTemp As String, lngOrderID As Long
    On Error GoTo EH
    sTemp = sCSVPath
    Do While InStr(sTemp, "\") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "\") + 1): Loop
    Do While InStr(sTemp, "/") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "/") + 1): Loop
    If LCase(Right$(sTemp, 4)) = ".csv" Then sJobName = Left$(sTemp, Len(sTemp) - 4) Else sJobName = sTemp
    
    If Not gbln_ConnectToDB() Then MsgBox "无法连接 CDM 数据库", vbCritical: Exit Sub
    lngOrderID = ImportCSV(sCSVPath, sJobName, sCustomerName, sMaterialName)
    If lngOrderID = -1 Then Exit Sub
    If lngOrderID <= 0 Then Exit Sub
    If bRunNest Then
        Frame.ShowProgressBox "自动化生产排版", "正在执行批量生产 + 排版 ..."
        DoEvents
        Call g_Make_Master(CStr(lngOrderID))
        Frame.CloseProgressBox
        If CBool(GetSetting("LICOM AlphaDOOR", "Nest Parameters", "Nest Completed", 0)) Then
            MsgBox "自动化生产排版完成！" & vbCrLf & "订单: " & sJobName, vbInformation
        Else
            MsgBox "生产排版可能未完全成功", vbExclamation
        End If
    Else
        MsgBox "CSV 导入完成！订单: " & sJobName, vbInformation
    End If
    Exit Sub
EH:
    Frame.CloseProgressBox
    MsgBox "错误: " & Err.Description, vbCritical
End Sub

Private Function ImportCSV(ByVal sCSVPath As String, ByVal sJobName As String, _
                           Optional ByVal sCustomerName As String = "自动化生产", _
                           Optional ByVal sDefaultMaterial As String = "开料机3000mm") As Long
    '
    Dim lngOrderID As Long, lngRow As Long, lngOK As Long
    Dim sLine As String, vF As Variant, iFile As Integer
    Dim lngCustID As Long
    Dim sTp As String, w As Double, h As Double, q As Long
    Dim sMat As String, sCu As String, sRf As String, sRm As String
    Dim sC1 As String, sC2 As String, sGrp As String
    Dim lngStyleNum As Long, sUsrStyle As String
    Dim lngFail As Long, lngIns As Long
    If sCustomerName = "" Then sCustomerName = "自动化生产"
    If sDefaultMaterial = "" Then sDefaultMaterial = "开料机3000mm"
    
    On Error GoTo EH
    
    ' 创建订单
    lngCustID = glng_EnsureCustomer(sCustomerName)
    lngOrderID = glng_CreateOrder(sJobName, lngCustID)
    If lngOrderID <= 0 Then GoTo CleanUp
    
    ' 检查文件是否存在
    If Dir(sCSVPath) = "" Then
        MsgBox "文件不存在:" & vbCrLf & sCSVPath, vbExclamation, "自动化生产排版"
        ImportCSV = 0
        Exit Function
    End If
    
    ' 读取 CSV
    iFile = FreeFile
    Open sCSVPath For Input As #iFile
    If Not EOF(iFile) Then Line Input #iFile, sLine
    
    Do While Not EOF(iFile)
        Line Input #iFile, sLine: lngRow = lngRow + 1
        sLine = Trim$(sLine): If sLine = "" Then GoTo NextLine
        vF = SplitCSVLine(sLine)
        If UBound(vF) < 3 Then GoTo NextLine
        
        sGrp = Trim$(GetF(vF, 4, "")): sTp = Trim$(GetF(vF, 0, "")): w = Val(GetF(vF, 1, "0"))
        h = Val(GetF(vF, 2, "0")): q = Val(GetF(vF, 3, "1"))
        sMat = Trim$(GetF(vF, 12, "")): If sMat = "" Then sMat = sDefaultMaterial
        sCu = Trim$(GetF(vF, 5, "")): sRf = Trim$(GetF(vF, 9, ""))
        sRm = Trim$(GetF(vF, 11, "")): sC1 = Trim$(GetF(vF, 7, "")): sC2 = Trim$(GetF(vF, 8, ""))
        If w <= 0 Or h <= 0 Then GoTo NextLine
        If q <= 0 Then q = 1
        
        ' 确保门型和材料存在，获取门型实际 StyleNumber 和用户样式名
        lngStyleNum = glng_EnsureStyle(sTp, sUsrStyle)
        glng_EnsureMaterial sMat
        
        ' 插入明细（INSERT...SELECT 从 AD_DOOR_TYPES 复制用户样式参数含 UserValue_0~6）
        gdb_CDM.Execute "INSERT INTO AD_ORDER_DETAILS " & _
            "(OrderID,TypeName,StyleName,StyleNumber,Quantity,Width,Length," & _
            "Material,ProductionComment,CSV_CustomerName,CSV_OrderNumber,CSV_ItemNumber," & _
            "CustomField1,CustomField2,ComponentGrouping,CornerRadius,RotationMethod,RotationAngle," & _
            "IgnoreOuterGeometry,ByPassNest,UserVariableString,UserDescriptionString," & _
            "UserValue_0,UserValue_1,UserValue_2,UserValue_3,UserValue_4,UserValue_5,UserValue_6) " & _
            "SELECT " & lngOrderID & ",'" & gs_FixSQL(sTp) & "','" & gs_FixSQL(sUsrStyle) & "'," & lngStyleNum & "," & _
            q & "," & w & "," & h & ",'" & gs_FixSQL(sMat) & "'," & _
            "'" & gs_FixSQL(sRm) & "','" & gs_FixSQL(sCu) & "'," & _
            "'" & gs_FixSQL(sRf) & "','" & gs_FixSQL(sGrp) & "'," & _
            "'" & gs_FixSQL(sC1) & "','" & gs_FixSQL(sC2) & "'," & _
            Val(sGrp) & "," & _
            "dt.CornerRadius,dt.RotationMethod,dt.RotationAngle," & _
            "dt.IgnoreOuterGeometry,dt.ByPassNest," & _
            "dt.UserVariableString,dt.UserDescriptionString," & _
            "dt.UserValue_0,dt.UserValue_1,dt.UserValue_2,dt.UserValue_3,dt.UserValue_4,dt.UserValue_5,dt.UserValue_6 " & _
            "FROM AD_DOOR_TYPES dt WHERE dt.TypeID='" & gs_FixSQL(sTp) & "'", lngIns
        If lngIns > 0 Then lngOK = lngOK + 1 Else lngFail = lngFail + 1
NextLine:
    Loop
    Close #iFile
    
    If lngFail > 0 Then
        MsgBox "有 " & lngFail & " 行门板明细未能插入（门型参数缺失），已跳过。", _
               vbExclamation, "自动化生产排版"
    End If
    
    ImportCSV = lngOrderID
    GoTo CleanUp
    
EH:
    ImportCSV = 0
    MsgBox "第 " & lngRow & " 行错误: " & Err.Description, vbExclamation
CleanUp:
    Close #iFile
End Function

Private Function glng_EnsureCustomer(ByVal sName As String) As Long
    Dim rst As ADODB.Recordset, lngRet As Long
    Set rst = New ADODB.Recordset
    rst.Open "SELECT CustomerID FROM AD_CUSTOMERS WHERE Name='" & gs_FixSQL(sName) & "'", gdb_CDM, adOpenForwardOnly, adLockReadOnly
    If rst.BOF And rst.EOF Then
        rst.Close
        gdb_CDM.Execute "INSERT INTO AD_CUSTOMERS (Name) VALUES ('" & gs_FixSQL(sName) & "')", lngRet
        Set rst = gdb_CDM.Execute("SELECT @@IDENTITY AS NewID")
        glng_EnsureCustomer = rst.Fields("NewID")
        rst.Close
    Else
        glng_EnsureCustomer = rst.Fields("CustomerID")
        rst.Close
    End If
End Function

Private Function glng_CreateOrder(ByVal sName As String, ByVal lngCustID As Long) As Long
    Dim rst As ADODB.Recordset, lngRet As Long
    ' 检查是否存在相同订单名，存在则直接取消
    Set rst = New ADODB.Recordset
    rst.Open "SELECT OrderID FROM AD_ORDERS WHERE JobName='" & gs_FixSQL(sName) & "'", gdb_CDM, adOpenForwardOnly, adLockReadOnly
    If Not (rst.BOF And rst.EOF) Then
        rst.Close
        MsgBox "订单名已存在，导入已取消: " & sName, _
               vbExclamation, "自动化生产排版"
        glng_CreateOrder = -1   ' 取消
        Exit Function
    End If
    rst.Close
    gdb_CDM.Execute "INSERT INTO AD_ORDERS (JobName,CustomerID,OrderDate) VALUES ('" & gs_FixSQL(sName) & "'," & lngCustID & ",Date())", lngRet
    If lngRet > 0 Then Set rst = gdb_CDM.Execute("SELECT @@IDENTITY AS NewID"): glng_CreateOrder = rst.Fields("NewID"): rst.Close
End Function

Private Function glng_EnsureStyle(ByVal sName As String, ByRef sUserStyleName As String) As Long
    Dim rst As ADODB.Recordset, lngRet As Long
    If sName = "" Then
        glng_EnsureStyle = 900
        Exit Function
    End If
    Set rst = New ADODB.Recordset
    rst.Open "SELECT TypeID,UserStyle,UserStyleName FROM AD_DOOR_TYPES WHERE TypeID='" & gs_FixSQL(sName) & "'", gdb_CDM, adOpenForwardOnly, adLockReadOnly
    If rst.BOF And rst.EOF Then
        rst.Close
        gdb_CDM.Execute "INSERT INTO AD_DOOR_TYPES (TypeID,UserStyle,UserStyleName,Width,Length) VALUES ('" & gs_FixSQL(sName) & "',False,'',900,900)", lngRet
        glng_EnsureStyle = 900
    Else
        If CBool(rst.Fields("UserStyle")) Then
            glng_EnsureStyle = 930
            sUserStyleName = gvar_CheckNull(rst.Fields("UserStyleName"))
        Else
            glng_EnsureStyle = 900
        End If
        rst.Close
    End If
End Function

Private Sub glng_EnsureMaterial(ByVal sName As String)
    Dim rst As ADODB.Recordset, lngRet As Long
    If sName = "" Then Exit Sub
    Set rst = New ADODB.Recordset
    rst.Open "SELECT Name FROM AD_MATERIALS WHERE Name='" & gs_FixSQL(sName) & "'", gdb_CDM, adOpenForwardOnly, adLockReadOnly
    If rst.BOF And rst.EOF Then rst.Close: gdb_CDM.Execute "INSERT INTO AD_MATERIALS (Name,Thickness,SheetWidth,SheetLength) VALUES ('" & gs_FixSQL(sName) & "',18,1220,3000)", lngRet
    rst.Close
End Sub

' ============================================================================
' CSV 解析
' ============================================================================
Private Function SplitCSVLine(ByVal sLine As String) As Variant
    Dim v() As String, idx As Long, i As Long, f As String, bQ As Boolean
    Dim c As String
    ReDim v(0 To 20)
    For i = 1 To Len(sLine)
        c = Mid$(sLine, i, 1)
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
                If idx > UBound(v) Then ReDim Preserve v(0 To idx * 2)
                v(idx) = f
                idx = idx + 1
                f = ""
            Else
                f = f & c
            End If
        End If
    Next
    If idx > UBound(v) Then ReDim Preserve v(0 To idx)
    v(idx) = f
    ReDim Preserve v(0 To idx)
    SplitCSVLine = v
End Function

Private Function GetF(ByRef v As Variant, ByVal i As Integer, ByVal d As String) As String
    If i >= 0 And i <= UBound(v) Then GetF = v(i) Else GetF = d
End Function

