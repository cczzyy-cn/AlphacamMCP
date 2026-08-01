Option Explicit

' ==============================================================================
' 自动化生产排版 — 弹窗选CSV → 导入 → 批量生产 → 排版（CDM 工程内）
' ==============================================================================
' 安装: 在 CDM 工程的 VBA 编辑器中导入此文件
' 用法: AlphaCAM 菜单 → CCC功能 → 自动化生产排版
' 说明: 由 CCC 功能菜单触发，代码在 CDM 工程中执行以直接调用 g_Make_Master
' ==============================================================================

' API 文件选择对话框


' ============================================================================
' 主入口（CCC 功能菜单触发）
' ============================================================================
Public Sub AutoImportNest()
    MsgBox "AutoImportNest 已调用", vbInformation
    '
    Dim sCSVPath As String
    Dim sJobName As String
    Dim sTemp    As String
    
    On Error GoTo EH
    
    ' ── 1. 选择 CSV 文件 ──
    sCSVPath = ShowOpenFile("C:\Users\C\Desktop\2026优化表", "CSV 文件|*.csv|所有文件|*.*")
    If sCSVPath = "" Then Exit Sub
    
    ' 取文件名作为订单名
    sTemp = sCSVPath
    Do While InStr(sTemp, "\") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "\") + 1): Loop
    Do While InStr(sTemp, "/") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "/") + 1): Loop
    If LCase(Right$(sTemp, 4)) = ".csv" Then sJobName = Left$(sTemp, Len(sTemp) - 4) Else sJobName = sTemp
    
    ' ── 2. 连接 DB 并导入 CSV ──
    If Not gbln_ConnectToDB() Then
        MsgBox "无法连接 CDM 数据库", vbCritical: Exit Sub
    End If
    
    Dim lngOrderID As Long
    lngOrderID = ImportCSV(sCSVPath, sJobName)
    If lngOrderID <= 0 Then Exit Sub
    
    ' ── 3. 调用 g_Make_Master 批量生产+排版 ──
    Frame.ShowProgressBox "自动化生产排版", "正在执行批量生产 + 排版 ..."
    DoEvents
    Call g_Make_Master(CStr(lngOrderID))
    Frame.CloseProgressBox
    
    ' ── 4. 完成 ──
    Dim bOK As Boolean
    bOK = CBool(GetSetting("LICOM AlphaDOOR", "Nest Parameters", "Nest Completed", 0))
    If bOK Then
        MsgBox "自动化生产排版完成！" & vbCrLf & "订单: " & sJobName, vbInformation
    Else
        MsgBox "生产排版可能未完全成功，请检查 AlphaCAM 结果。", vbExclamation
    End If
    Exit Sub
    
EH:
    Frame.CloseProgressBox
    MsgBox "错误: " & Err.Description, vbCritical
End Sub


' ============================================================================
' 导入 CSV → 创建订单，返回 OrderID
' ============================================================================
Private Function ImportCSV(ByVal sCSVPath As String, ByVal sJobName As String) As Long
    '
    Dim lngOrderID As Long, lngRow As Long, lngOK As Long
    Dim sLine As String, vF As Variant, iFile As Integer
    
    On Error GoTo EH
    
    ' 创建订单
    lngOrderID = glng_CreateOrder(sJobName, 1)
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
        
        Dim sTp As String, w As Double, h As Double, q As Long
        Dim sMat As String, sCu As String, sRf As String, sRm As String
        Dim sC1 As String, sC2 As String
        
        sTp = Trim$(GetF(vF, 0, "")): w = Val(GetF(vF, 1, "0"))
        h = Val(GetF(vF, 2, "0")): q = Val(GetF(vF, 3, "1"))
        sMat = Trim$(GetF(vF, 12, "")): If sMat = "" Then sMat = "开料机3000mm"
        sCu = Trim$(GetF(vF, 5, "")): sRf = Trim$(GetF(vF, 9, ""))
        sRm = Trim$(GetF(vF, 11, "")): sC1 = Trim$(GetF(vF, 7, "")): sC2 = Trim$(GetF(vF, 8, ""))
        If w <= 0 Or h <= 0 Then GoTo NextLine
        If q <= 0 Then q = 1
        
        ' 确保门型和材料存在
        glng_EnsureStyle sTp
        glng_EnsureMaterial sMat
        
        ' 插入明细
        gdb_CDM.Execute "INSERT INTO AD_ORDER_DETAILS " & _
            "(OrderID,TypeName,StyleName,StyleNumber,Quantity,Width,Length," & _
            "Material,ProductionComment,CSV_CustomerName,CSV_OrderNumber,CSV_ItemNumber," & _
            "CustomField1,CustomField2) VALUES (" & _
            lngOrderID & ",'" & gs_FixSQL(sTp) & "','" & gs_FixSQL(sTp) & "',900," & _
            q & "," & w & "," & h & ",'" & gs_FixSQL(sMat) & "'," & _
            "'" & gs_FixSQL(sRm) & "','" & gs_FixSQL(sCu) & "'," & _
            "'" & gs_FixSQL(sRf) & "','" & gs_FixSQL(sRf) & "'," & _
            "'" & gs_FixSQL(sC1) & "','" & gs_FixSQL(sC2) & "')"
        lngOK = lngOK + 1
NextLine:
    Loop
    Close #iFile
    
    ImportCSV = lngOrderID
    GoTo CleanUp
    
EH:
    ImportCSV = 0
    MsgBox "第 " & lngRow & " 行错误: " & Err.Description, vbExclamation
CleanUp:
    Close #iFile
End Function

Private Function glng_CreateOrder(ByVal sName As String, ByVal lngCustID As Long) As Long
    Dim rst As ADODB.Recordset, lngRet As Long
    gdb_CDM.Execute "INSERT INTO AD_ORDERS (JobName,CustomerID,OrderDate) VALUES ('" & gs_FixSQL(sName) & "'," & lngCustID & ",Date())", lngRet
    If lngRet > 0 Then Set rst = gdb_CDM.Execute("SELECT @@IDENTITY AS NewID"): glng_CreateOrder = rst.Fields("NewID"): rst.Close
End Function

Private Sub glng_EnsureStyle(ByVal sName As String)
    Dim rst As ADODB.Recordset, lngRet As Long
    If sName = "" Then Exit Sub
    Set rst = New ADODB.Recordset
    rst.Open "SELECT PK FROM AD_DOOR_TYPES WHERE TypeID='" & gs_FixSQL(sName) & "'", gdb_CDM, adOpenForwardOnly, adLockReadOnly
    If rst.BOF And rst.EOF Then rst.Close: gdb_CDM.Execute "INSERT INTO AD_DOOR_TYPES (TypeID,StyleNumber) VALUES ('" & gs_FixSQL(sName) & "',900)", lngRet
    rst.Close
End Sub

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
                v(idx) = f
                idx = idx + 1
                f = ""
            Else
                f = f & c
            End If
    Next
    v(idx) = f
    ReDim Preserve v(0 To idx)
    SplitCSVLine = v
End Function

Private Function GetF(ByRef v As Variant, ByVal i As Integer, ByVal d As String) As String
    If i >= 0 And i <= UBound(v) Then GetF = v(i) Else GetF = d
End Function

' ============================================================================
' 文件选择对话框
' ============================================================================
Private Function ShowOpenFile(ByVal sDir As String, ByVal sFilter As String) As String
    Dim sLast As String
    Dim sInput As String
    
    ' 读取记忆的完整路径
    sLast = GetSetting("CCC", "AutoImportNest", "LastPath", sDir & "\7-10中林SPC婷兰灰.csv")
    
    ' 单个输入框：完整路径（带记忆）
    sInput = InputBox("请输入 CSV 文件完整路径:" & vbCrLf & vbCrLf & _
                      "可以只输入文件名（使用记忆的目录）:" & vbCrLf & _
                      "  " & sDir & vbCrLf & vbCrLf & _
                      "或输入完整路径:", _
                      "自动化生产排版", sLast)
    If sInput = "" Then Exit Function
    
    ' 如果只输入文件名，拼接记忆目录
    If InStr(sInput, "\") = 0 And InStr(sInput, "/") = 0 Then
        ' 从记忆路径提取目录
        Dim sFolder As String
        sFolder = Left$(sLast, InStrRev(sLast, "\"))
        If sFolder = "" Then sFolder = sDir & "\"
        sInput = sFolder & sInput
    End If
    
    ' 保存记忆
    SaveSetting "CCC", "AutoImportNest", "LastPath", sInput
    
    ShowOpenFile = sInput
End Function
