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
    ' 菜单入口：弹出 frmAutoNest 窗体（非模态，不阻塞 AlphaCAM）
    On Error GoTo EH
    frmAutoNest.Show vbModeless
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
                                  ByVal bRunNest As Boolean, _
                                  Optional ByVal bOverwrite As Boolean = False)
    ' 由 frmAutoNest 窗体调用的带参入口
    Dim sJobName As String, sTemp As String, lngOrderID As Long
    On Error GoTo EH
    sTemp = sCSVPath
    Do While InStr(sTemp, "\") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "\") + 1): Loop
    Do While InStr(sTemp, "/") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "/") + 1): Loop
    If LCase(Right$(sTemp, 4)) = ".csv" Then sJobName = Left$(sTemp, Len(sTemp) - 4) Else sJobName = sTemp
    
    If Not gbln_ConnectToDB() Then MsgBox "无法连接 CDM 数据库", vbCritical: Exit Sub
    lngOrderID = ImportCSV(sCSVPath, sJobName, sCustomerName, sMaterialName, bOverwrite)
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
                           Optional ByVal sDefaultMaterial As String = "开料机3000mm", _
                           Optional ByVal bOverwrite As Boolean = False) As Long
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
    lngOrderID = glng_CreateOrder(sJobName, lngCustID, bOverwrite)
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

Private Function glng_CreateOrder(ByVal sName As String, ByVal lngCustID As Long, _
                                  Optional ByVal bOverwrite As Boolean = False) As Long
    Dim rst As ADODB.Recordset, lngRet As Long
    Dim lngOldID As Long
    ' 检查是否存在相同订单名
    Set rst = New ADODB.Recordset
    rst.Open "SELECT OrderID FROM AD_ORDERS WHERE JobName='" & gs_FixSQL(sName) & "'", gdb_CDM, adOpenForwardOnly, adLockReadOnly
    If Not (rst.BOF And rst.EOF) Then
        lngOldID = rst.Fields("OrderID")
        rst.Close
        If bOverwrite Then
            ' 强制覆盖：删除原订单所有相关数据（明细/报表），再删订单
            On Error Resume Next
            gdb_CDM.Execute "DELETE FROM AD_ORDER_DETAILS WHERE OrderID=" & lngOldID
            gdb_CDM.Execute "DELETE FROM AD_REPORT_DATA WHERE OrderID=" & lngOldID
            gdb_CDM.Execute "DELETE FROM AD_ORDERS WHERE OrderID=" & lngOldID
            On Error GoTo 0
        Else
            MsgBox "订单名已存在，导入已取消: " & sName, _
                   vbExclamation, "自动化生产排版"
            glng_CreateOrder = -1   ' 取消
            Exit Function
        End If
    Else
        rst.Close
    End If
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

' ============================================================================
' 重新生成门板标签 EMF（手动移动门板后使用）
' 用法: 打开排版后的嵌套图纸（含移动后的门板位置）→ 运行本宏
' 效果: 按当前图纸状态逐件重新生成 <JobName>_<材料>_<板名>_<件号>.emf，
'       覆盖原文件（AD_REPORT_DATA.PressDoorImage 路径不变，内容已更新）
' ============================================================================
Public Sub g_RegenDoorLabelEMFs()
    On Error GoTo EH
    Dim Material As CMaterial
    Dim p As Path
    Dim sMat As String
    Dim sName As String, iPos As Long
    Dim rst As ADODB.Recordset
    Dim sPath As String, sBase As String
    
    ' 1. 初始化选项（路径/报表配置与排版时一致）
    Set clsOptions = New COptions
    strCTX = clsOptions.CTXFile
    
    ' 2. 从刀路属性恢复订单名（m_SetAttributes 写入 DEF_ATT_JOB_NAME）
    gstr_JobName = ""
    For Each p In ActiveDrawing.ToolPaths
        If p.Attribute(DEF_ATT_JOB_NAME) <> "" Then
            gstr_JobName = p.Attribute(DEF_ATT_JOB_NAME)
            Exit For
        End If
    Next
    If gstr_JobName = "" Then
        MsgBox "无法从刀路属性恢复订单名（DEF_ATT_JOB_NAME）。" & vbCrLf & _
               "请确认打开的是排版后的嵌套图纸（含 AlphaDOOR 刀路属性）。", _
               vbExclamation, "自动化生产排版"
        Exit Sub
    End If
    
    ' 3. 恢复材料名（注意：嵌套板 MaterialName 是 SheetName 配置名如 "Admin"，非材料名）
    '    优先级：图纸名解析 <JobName>_<材料>.ard → 数据库 PressDoorImage 路径解析 → 图纸属性
    sMat = ""
    sName = ActiveDrawing.Name
    If InStr(sName, ".") > 0 Then sName = Left$(sName, InStr(sName, ".") - 1)
    iPos = InStr(sName, gstr_JobName)
    If iPos > 0 Then sMat = Mid$(sName, iPos + Len(gstr_JobName) + 1)
    If sMat = "" Then
        ' 数据库回退：PressDoorImage 路径含真实材料名
        On Error Resume Next
        If gbln_ConnectToDB() Then
            Set rst = gdb_CDM.Execute("SELECT TOP 1 PressDoorImage FROM AD_REPORT_DATA WHERE PressDoorImage <> '' AND INSTR(PressDoorImage, '" & gs_FixSQL(gstr_JobName) & "') > 0")
            If Not rst Is Nothing Then
                If Not rst.EOF Then
                    sPath = rst.Fields(0)
                    sBase = Mid$(sPath, InStrRev(sPath, "\") + 1)
                    If InStr(sBase, ".") > 0 Then sBase = Left$(sBase, InStr(sBase, ".") - 1)
                    iPos = InStr(sBase, gstr_JobName)
                    If iPos > 0 Then
                        sMat = Mid$(sBase, iPos + Len(gstr_JobName) + 1)
                        If InStr(sMat, "_") > 0 Then sMat = Left$(sMat, InStr(sMat, "_") - 1)
                    End If
                End If
                rst.Close
            End If
        End If
        On Error GoTo 0
    End If
    If sMat = "" Then sMat = ActiveDrawing.Attribute(DEF_ATT_MATERIAL_NAME)
    If sMat = "" Then
        MsgBox "无法恢复材料名（图纸名/数据库/图纸属性均为空）。", _
               vbExclamation, "自动化生产排版"
        Exit Sub
    End If
    
    ' 4. 数据库连接 + 排版区域集合（m_CreateAlphaCAMDrawingsOfSheets 依赖）
    If Not gbln_ConnectToDB() Then
        MsgBox "无法连接 CDM 数据库", vbCritical: Exit Sub
    End If
    m_PopulateNestingZones
    
    ' 5. 构造材料对象并重新生成 EMF
    Set Material = New CMaterial
    Material.MaterialName = sMat
    m_CreateAlphaCAMDrawingsOfSheets Material
    
    MsgBox "门板标签 EMF 已按当前门板位置重新生成" & vbCrLf & vbCrLf & _
           "订单: " & gstr_JobName & vbCrLf & _
           "材料: " & sMat, vbInformation, "自动化生产排版"
    Exit Sub
EH:
    MsgBox "错误: " & Err.Description, vbCritical, "自动化生产排版"
End Sub
