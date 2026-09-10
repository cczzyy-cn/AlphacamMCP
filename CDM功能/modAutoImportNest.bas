' MCP-INSTALL-TEST: 2026-07-30 测试注释
' ==============================================================================
' 版本: v1.8 (2026-09-10) — 修复屏幕刷新泄漏（窗口停在临时档案名的根源）+ 备份/临时文件清理
'   本次变更（承 v1.7）:
'     [P0] 配合 Make.bas:3991/3992 恢复 ActiveDrawing.ScreenUpdating 与
'          Frame.ProjectBarUpdating，并在重生成的成功与失败路径都兜底恢复刷新 + Redraw。
'          m_CreateAlphaCAMDrawingsOfSheets 会把两者置 False 且原恢复语句被注释，
'          导致每次生产/重生成后 AlphaCAM 不再重绘 —— 窗口标题与画面就停在
'          regen_<订单>_<Timer> 这个临时副本名上（实际活动文档已是真档案）。
'     [P1] 5.3b 备份目录改为先 Kill 再 RmDir：RmDir 对非空目录必定失败，
'          原实现从未删成功，ProgramData 累积了 28 个 regen_backup_*（222 个旧 EMF）。
'     [P1] 临时嵌套 ard 改在 App.New 关档之后再补删一次，结果写入 CDM_Import.log。
'   ---- 以下为 v1.7 变更记录 ----
' 版本: v1.7 (2026-09-10) — 健壮性修复：CSV 导入真正事务化、重生成失败现场还原、清死代码
'   本次变更（承 v1.6）:
'     [P0] ImportCSV 全流程纳入 BeginTrans/CommitTrans/RollbackTrans：
'          覆盖模式下「先删旧订单、再插新明细」中途失败时不再丢旧数据
'     [P0] 重生成标签失败时：还原旧 EMF + 删除 regen_* 临时文件
'          + 重新打开用户原档案（原实现失败后把用户留在空白图/临时副本里）
'     [P1] 重生成前新增「未保存修改」确认提示（未保存的手动移动不会写回原图）
'     [P1] CSV 文件存在性检查上移到创建订单之前（不再产生空订单记录）
'     [P1] 删除死代码：sDefaultMaterial 形参、glng_EnsureMaterial 过程
'     [P2] Sleep 声明补 PtrSafe（与 frmAutoNest 保持一致）
'     [P2] m_WaitForLabelEMFs 期望值改为「板数 + 总件数」：原实现只传总件数，
'          而目录计数会把整板图 <Job>_<Mat>_<板名>.emf 一并计入，导致提前返回
'   历史:
'     v1.6 (2026-09-02) 重生成标签移除对主图的 Name/Kill（闪退源），不重存主图、结尾重开真档案
'       [P0] 旧EMF 改为先备份到目录、成功后删除、失败回滚
'       [P0] 数据库同步改用事务 BeginTrans/CommitTrans/RollbackTrans
'       [P0] 去掉 Exit For 漏更，多路径板件全更
'       [P1] 标签导出改由 Make.m_ExportDoorLabelEMFs 统一生成
'       [v1.3] 重生成传临时巢套路径，绝不 SaveAs 覆盖用户主图，保留加工道次
'       [v1.5] 删除对主图 .bak 驻留/还原（Name/Kill，闪退根源）；结尾重开真档案
'   日期: 2026-09-10
' ==============================================================================
Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ==============================================================================
' 自动化生产排版 — 弹窗选CSV → 导入 → 批量生产 → 排版（CDM 工程内）
' ==============================================================================
' 安装: 在 CDM 工程的 VBA 编辑器中导入此文件
' 用法: AlphaCAM 菜单 → CCC功能 → 自动化生产排版
' 说明: 由 CCC 功能菜单触发，代码在 CDM 工程中执行以直接调用 g_Make_Master
'
' v1.1 增强（错误与提示处理）:
'   - 错误日志落盘（CDM_Import.log，含时间/步骤/错误号/描述）
'   - 错误提示带步骤名 + Err.Number
'   - CSV 导入使用事务，失败回滚本次订单明细/客户/门型/材料
'   - 失败行明细列表，区分失败原因
'   - 覆盖删除失败不再静默吞错
'   - 返回码用常量(ORDER_CANCEL/ORDER_FAIL)替代魔法数字
'   - Nest Completed 残留键清除，避免误判
' ==============================================================================

' ============================================================================
' 常量 / 返回码
' ============================================================================
Private Const ORDER_CANCEL As Long = -1      ' 取消导入（订单名冲突且不覆盖）
Private Const ORDER_FAIL As Long = 0         ' 导入失败
Private Const LOG_FILE_NAME As String = "CDM_Import.log"

' ============================================================================
' 日志（落盘到程序共用数据目录）
' ============================================================================
Private Sub m_Log(ByVal sMsg As String)
    Dim iFileNo As Integer
    On Error Resume Next
    iFileNo = FreeFile
    Open gs_GetCommonAppDataDir & LOG_FILE_NAME For Append As #iFileNo
    Print #iFileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & sMsg
    Close #iFileNo
    On Error GoTo 0
End Sub

Private Sub m_LogError(ByVal lngNum As Long, ByVal sStep As String, ByVal sDesc As String)
    m_Log "[ERROR] 步骤=" & IIf(sStep = "", "?", sStep) & " 编号=" & CStr(lngNum) & _
          " 描述=" & sDesc
End Sub

' ============================================================================
' 主入口（CCC 功能菜单触发）
' ============================================================================
Public Sub AutoImportNest()
    ' 菜单入口：弹出 frmAutoNest 窗体（非模态，不阻塞 AlphaCAM）
    On Error GoTo EH
    frmAutoNest.Show vbModeless
    Exit Sub
EH:
    m_LogError Err.Number, "AutoImportNest", Err.Description
    MsgBox "错误: " & Err.Description, vbCritical
End Sub

' ============================================================================
' 导入 CSV → 创建订单，返回 OrderID
' ============================================================================
' 注意: 无材料形参 —— 材料逐行取自 CSV 第 12 列（0 基），且必须已存在于
'       AD_MATERIALS 表；原 sMaterialName 形参只用于填充一个从未被使用的
'       默认值（死代码），v1.7 已删除。
Public Sub AutoImportNestWithParams(ByVal sCSVPath As String, _
                                    ByVal sCustomerName As String, _
                                    ByVal bRunNest As Boolean, _
                                    Optional ByVal bOverwrite As Boolean = False)
    ' 由 frmAutoNest 窗体调用的带参入口
    Dim sJobName As String, sTemp As String, lngOrderID As Long
    Dim sStep As String
    sStep = "解析CSV"
    On Error GoTo EH
    m_Log "开始导入 CSV=" & sCSVPath & " 客户=" & sCustomerName & _
          " 运行排版=" & CStr(bRunNest) & " 覆盖=" & CStr(bOverwrite)

    sTemp = sCSVPath
    Do While InStr(sTemp, "\") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "\") + 1): Loop
    Do While InStr(sTemp, "/") > 0: sTemp = Mid$(sTemp, InStr(sTemp, "/") + 1): Loop
    If LCase(Right$(sTemp, 4)) = ".csv" Then sJobName = Left$(sTemp, Len(sTemp) - 4) Else sJobName = sTemp

    sStep = "连接数据库"
    If Not gbln_ConnectToDB() Then
        m_LogError 0, sStep, "无法连接CDM数据库"
        MsgBox "无法连接 CDM 数据库", vbCritical
        Exit Sub
    End If

    sStep = "导入CSV"
    lngOrderID = ImportCSV(sCSVPath, sJobName, sCustomerName, bOverwrite)
    m_Log "导入CSV返回 OrderID=" & CStr(lngOrderID)
    If lngOrderID = ORDER_CANCEL Then Exit Sub     ' 取消：内部已提示
    If lngOrderID = ORDER_FAIL Then Exit Sub       ' 失败：ImportCSV 内已提示+日志

    If bRunNest Then
        sStep = "批量生产+排版"
        Frame.ShowProgressBox "自动化生产排版", "正在执行批量生产 + 排版 ..."
        DoEvents
        ' 清除历史残留的 Nest Completed 键，避免上一订单残留导致误判
        On Error Resume Next
        DeleteSetting "LICOM AlphaDOOR", "Nest Parameters", "Nest Completed"
        On Error GoTo EH
        Call g_Make_Master(CStr(lngOrderID))
        Frame.CloseProgressBox
        If CBool(GetSetting("LICOM AlphaDOOR", "Nest Parameters", "Nest Completed", 0)) Then
            m_Log "生产排版完成: 订单=" & sJobName & " OrderID=" & CStr(lngOrderID)
            MsgBox "自动化生产排版完成！" & vbCrLf & "订单: " & sJobName, vbInformation
        Else
            m_LogError 0, sStep, "生产排版可能未完全成功: 订单=" & sJobName & " OrderID=" & CStr(lngOrderID)
            MsgBox "生产排版可能未完全成功", vbExclamation
        End If
    Else
        m_Log "CSV仅导入完成: 订单=" & sJobName
        MsgBox "CSV 导入完成！订单: " & sJobName, vbInformation
    End If
    Exit Sub
EH:
    Frame.CloseProgressBox
    m_LogError Err.Number, sStep, Err.Description
    MsgBox "步骤[" & sStep & "] 错误: " & Err.Description, vbCritical
End Sub

' ============================================================================
' CSV 导入（事务包裹 + 失败明细 + 返回码常量 + 日志）
' 材料: 逐行取自 CSV 第 12 列（0 基），必须已存在于 AD_MATERIALS 表，
'       缺失即整体失败并回滚（本模块只校验、不自动建档）
' ============================================================================
Private Function ImportCSV(ByVal sCSVPath As String, ByVal sJobName As String, _
                           Optional ByVal sCustomerName As String = "自动化生产", _
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
    Dim sStep As String
    Dim colFailed As Collection
    Dim iMsg As Long
    Dim lngErrNum As Long, sErrDesc As String, sInsSQL As String
    Dim blnInTrans As Boolean

    If sCustomerName = "" Then sCustomerName = "自动化生产"
    Set colFailed = New Collection
    On Error GoTo EH

    ' 1. 文件存在性检查必须在创建订单之前，否则路径写错会留下空订单 + 客户记录
    sStep = "检查CSV文件"
    If Dir(sCSVPath) = "" Then
        m_LogError 0, sStep, "文件不存在: " & sCSVPath
        MsgBox "文件不存在:" & vbCrLf & sCSVPath, vbExclamation, "自动化生产排版"
        ImportCSV = ORDER_FAIL
        GoTo CleanUp
    End If

    ' 2. 开启事务：客户/订单/门型/明细任一步失败都整体回滚。
    '    对覆盖模式尤其关键 —— 旧订单的 DELETE 也在同一事务内，
    '    中途失败时旧数据随回滚恢复，不会出现「旧的删了、新的没插全」。
    sStep = "开启数据库事务"
    gdb_CDM.BeginTrans
    blnInTrans = True

    sStep = "创建订单"
    lngCustID = glng_EnsureCustomer(sCustomerName)
    lngOrderID = glng_CreateOrder(sJobName, lngCustID, bOverwrite)
    If lngOrderID = ORDER_CANCEL Then
        ImportCSV = ORDER_CANCEL
        GoTo CleanUp
    End If
    If lngOrderID = ORDER_FAIL Then
        ImportCSV = ORDER_FAIL
        GoTo CleanUp
    End If

    ' 读取 CSV
    sStep = "读取CSV"
    iFile = FreeFile
    Open sCSVPath For Input As #iFile
    If Not EOF(iFile) Then Line Input #iFile, sLine   ' 表头

    sStep = "导入明细"
    Do While Not EOF(iFile)
        Line Input #iFile, sLine: lngRow = lngRow + 1
        sLine = Trim$(sLine): If sLine = "" Then GoTo NextLine
        vF = SplitCSVLine(sLine)
        If UBound(vF) < 3 Then GoTo NextLine

        sGrp = Trim$(GetF(vF, 4, "")): sTp = Trim$(GetF(vF, 0, "")): w = Val(GetF(vF, 1, "0"))
        h = Val(GetF(vF, 2, "0")): q = Val(GetF(vF, 3, "1"))
        sMat = Trim$(GetF(vF, 12, ""))
        sCu = Trim$(GetF(vF, 5, "")): sRf = Trim$(GetF(vF, 9, ""))
        sRm = Trim$(GetF(vF, 11, "")): sC1 = Trim$(GetF(vF, 7, "")): sC2 = Trim$(GetF(vF, 8, ""))
        If w <= 0 Or h <= 0 Then GoTo NextLine
        If q <= 0 Then q = 1

        ' 确保门型和材料存在，获取门型实际 StyleNumber 和用户样式名
        lngStyleNum = glng_EnsureStyle(sTp, sUsrStyle)
        ' 校验材料必须在数据库 AD_MATERIALS 存在，否则提示并结束导入
        If Not m_CheckMaterialExists(sMat) Then
            m_LogError 0, sStep, "材料[" & IIf(sMat = "", "无", sMat) & "]未在数据库AD_MATERIALS定义"
            ' 先回滚再弹窗：避免用户阅读提示期间事务一直占着锁
            On Error Resume Next
            If blnInTrans Then gdb_CDM.RollbackTrans
            On Error GoTo EH
            blnInTrans = False
            MsgBox "材料 [" & IIf(sMat = "", "无", sMat) & "] 不在数据库 AD_MATERIALS 表，无法导入。" & vbCrLf & "请先在材料库添加该材料后重试。", vbCritical, "自动化生产排版"
            ImportCSV = ORDER_FAIL
            GoTo CleanUp
        End If

        ' 插入明细（INSERT...SELECT 从 AD_DOOR_TYPES 复制用户样式参数含 UserValue_0~6）
        sInsSQL = "INSERT INTO AD_ORDER_DETAILS " & _
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
            "FROM AD_DOOR_TYPES dt WHERE dt.TypeID='" & gs_FixSQL(sTp) & "'"
        gdb_CDM.Execute sInsSQL, lngIns
        If lngIns > 0 Then
            lngOK = lngOK + 1
        Else
            ' 区分失败原因：INSERT...SELECT 无行 = 门型在 AD_DOOR_TYPES 无参数
            lngFail = lngFail + 1
            colFailed.add "第" & CStr(lngRow) & "行: 门型[" & sTp & "] 尺寸[" & CStr(w) & "x" & CStr(h) & "]x" & CStr(q) & " 参数缺失"
        End If
NextLine:
    Loop
    Close #iFile
    iFile = 0

    m_Log "导入明细完成: 成功=" & CStr(lngOK) & " 失败=" & CStr(lngFail) & _
          " (文件=" & sCSVPath & ")"

    ' 3. 明细已全部写入（失败行按“容忍部分失败”策略跳过）→ 提交事务
    sStep = "提交数据库事务"
    gdb_CDM.CommitTrans
    blnInTrans = False

    If lngFail > 0 Then
        Dim sMsg As String
        sMsg = "有 " & lngFail & " 行门板明细未能插入，已跳过。"
        If colFailed.Count > 0 Then
            sMsg = sMsg & vbCrLf & "失败明细(前10条):"
            For iMsg = 1 To colFailed.Count
                If iMsg > 10 Then Exit For
                sMsg = sMsg & vbCrLf & "  " & colFailed(iMsg)
            Next iMsg
            If colFailed.Count > 10 Then sMsg = sMsg & vbCrLf & "  ...共" & CStr(colFailed.Count) & "条(详见日志)"
        End If
        MsgBox sMsg, vbExclamation, "自动化生产排版"
    End If

    ImportCSV = lngOrderID
    GoTo CleanUp

EH:
    lngErrNum = Err.Number
    sErrDesc = Err.Description
    ' 先回滚再记录/提示：出错后数据库不留半截数据（含覆盖模式里已删掉的旧订单）
    On Error Resume Next
    If blnInTrans Then gdb_CDM.RollbackTrans
    blnInTrans = False
    m_LogError lngErrNum, sStep, "第" & CStr(lngRow) & "行: " & sErrDesc & " 错误号=" & CStr(lngErrNum) & " (文件=" & sCSVPath & ")"
    m_Log "  [SQL] " & sInsSQL
    MsgBox "步骤[" & sStep & "] 第" & CStr(lngRow) & "行 错误 " & CStr(lngErrNum) & ": " & sErrDesc & vbCrLf & vbCrLf & "SQL: " & Left$(sInsSQL, 400), vbCritical, "自动化生产排版"
    ImportCSV = ORDER_FAIL
CleanUp:
    ' 幂等收尾：未提交则回滚（涵盖提前 GoTo CleanUp 的取消/材料缺失分支），
    ' 并确保 CSV 文件句柄一定关闭
    On Error Resume Next
    If blnInTrans Then gdb_CDM.RollbackTrans
    blnInTrans = False
    If iFile <> 0 Then Close #iFile
    On Error GoTo 0
End Function

' ============================================================================
' 保证客户存在，返回 CustomerID
' ============================================================================
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

' ============================================================================
' 创建订单，返回 OrderID；重名：覆盖 bOverwrite 或取消返回 ORDER_CANCEL
' ============================================================================
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
            ' 强制覆盖：删除原订单所有相关数据（明细/报表），再删订单。
            ' 不再用 On Error Resume Next 吞错：任一删除失败会向上抛出，
            ' 由外层 ImportCSV 的事务回滚并提示，避免残留脏数据。
            m_Log "覆盖旧订单 OrderID=" & CStr(lngOldID) & " 名称=" & sName
            gdb_CDM.Execute "DELETE FROM AD_ORDER_DETAILS WHERE OrderID=" & lngOldID, lngRet
            gdb_CDM.Execute "DELETE FROM AD_REPORT_DATA WHERE OrderID=" & lngOldID, lngRet
            gdb_CDM.Execute "DELETE FROM AD_ORDERS WHERE OrderID=" & lngOldID, lngRet
        Else
            m_Log "订单名已存在，导入取消: " & sName
            MsgBox "订单名已存在，导入已取消: " & sName, _
                   vbExclamation, "自动化生产排版"
            glng_CreateOrder = ORDER_CANCEL   ' 取消
            Exit Function
        End If
    Else
        rst.Close
    End If
    gdb_CDM.Execute "INSERT INTO AD_ORDERS (JobName,CustomerID,OrderDate) VALUES ('" & gs_FixSQL(sName) & "'," & lngCustID & ",Date())", lngRet
    If lngRet > 0 Then Set rst = gdb_CDM.Execute("SELECT @@IDENTITY AS NewID"): glng_CreateOrder = rst.Fields("NewID"): rst.Close
End Function

' ============================================================================
' 保证门型存在，返回 StyleNumber（900 标准 / 930 用户样式）并带回用户样式名
' ============================================================================
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

' ============================================================================
' 校验材料是否存在于 AD_MATERIALS 表（只检查，不自动创建）
' v1.7: 原 glng_EnsureMaterial（自动建档 18mm/1220x3000）为死代码，已删除 ——
'       材料缺失一律报错并要求先在材料库添加，避免静默建出错规格的材料
' ============================================================================
Private Function m_CheckMaterialExists(ByVal sName As String) As Boolean
    Dim rst As ADODB.Recordset
    If sName = "" Then
        m_CheckMaterialExists = False
        Exit Function
    End If
    Set rst = New ADODB.Recordset
    rst.Open "SELECT Name FROM AD_MATERIALS WHERE Name='" & gs_FixSQL(sName) & "'", gdb_CDM, adOpenForwardOnly, adLockReadOnly
    m_CheckMaterialExists = Not (rst.BOF And rst.EOF)
    rst.Close
End Function

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
' v1.7: 开头先提示保存（未保存的手动移动不会写回原图，只用于出标签）；
'       失败时自动 回滚数据库 + 还原旧 EMF + 清临时文件 + 重开原档案
' ============================================================================
' 轮询等待标签 EMF 落盘完成（最多 60 x 100ms ≈ 6s）
' m_Create 本身是同步的，但窗口/文件写入是异步落地的，所以按文件轮询。
' 重要：本函数用 Dir 统计 <Job>_<Mat>_*.emf，该通配同时匹配
'   整板图 <Job>_<Mat>_<板名>.emf          （每板 1 张，Make.m_Create...:3941）
'   逐件图 <Job>_<Mat>_<板名>_<件号>.emf   （每板 = 该板 SH.Parts.Count 张）
' 因此调用方传入的 lngExpected 必须是「板数 + 总件数」；只传总件数会因
' 整板图被一并计入而提前满足条件（v1.7 修正，v1.6 及以前存在此缺陷）。
' ============================================================================
Private Function m_WaitForLabelEMFs(ByVal sJob As String, ByVal sMat As String, ByVal lngExpected As Long) As Boolean
    Dim i As Long, n As Long, sF As String
    Dim sDir As String
    sDir = gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE & DEF_BACKSLASH & sJob & "_" & sMat & "_"
    For i = 1 To 60
        DoEvents
        n = 0
        sF = Dir$(sDir & "*" & DEF_EXTENSION_EMF)
        Do While sF <> ""
            n = n + 1
            sF = Dir$
        Loop
        If n >= lngExpected Then m_WaitForLabelEMFs = True: Exit Function
        Sleep 100
    Next i
    m_WaitForLabelEMFs = (n >= lngExpected)
End Function

Public Sub g_RegenDoorLabelEMFs()
    On Error GoTo EH
    Dim Material As CMaterial
    Dim P As Path
    Dim sMat As String
    Dim sName As String, iPos As Long
    Dim rst As ADODB.Recordset
    Dim sPath As String, sBase As String
    Dim Ni As NestInformation
    Dim Nsh As NestSheet, Npi As NestPartInstance
    Dim rst2 As ADODB.Recordset
    Dim lngOrderID2 As Long, lngCnt As Long, lngDetail As Long
    Dim sDetailList As String, sImgPath As String, sDel As String
    Dim psTemp As Paths, lngBG As Long
    Dim blnGradW As Boolean, blnGradS As Boolean
    Dim dMinX As Double, dMinY As Double, dMaxX As Double, dMaxY As Double
    Dim colGeoColors As New Collection, colTPColors As New Collection
    Dim iClr As Long
    Dim psHatch As Paths, pOut As Path
    Dim sUserARD As String, sTmpBak As String
    Dim sNestOverride As String, sScratchFile As String
    Dim FSO2 As Scripting.FileSystemObject
    Dim sBakDir As String, blnInTrans As Boolean
    Dim lngExpectedDoors As Long, lngSheetCount As Long
    Dim blnScratchOpened As Boolean

    ' 0. 检查当前图纸是否为排版档案（ARD 嵌套图）
    '    注意: VBA 的 Or 不短路，Ni Is Nothing Or Ni.Sheets.Count 会在
    '    Ni 为 Nothing 时仍求值 Ni.Sheets 而报错 91，必须分开判断;
    '    Set 后用 On Error GoTo EH 恢复主错误处理（不能用 GoTo 0 禁用）
    Set Ni = Nothing
    On Error Resume Next
    Set Ni = ActiveDrawing.GetNestInformation
    On Error GoTo EH
    If Ni Is Nothing Then
        MsgBox "当前图纸不是排版档案！" & vbCrLf & vbCrLf & _
               "请先在 AlphaCAM 中打开排版后的嵌套图纸（ARD 文件，" & vbCrLf & _
               "如 <订单>_<材料>.ard），再点“重新生成标签”。", _
               vbExclamation, "自动化生产排版"
        Exit Sub
    End If
    If Ni.Sheets.Count = 0 Then
        MsgBox "当前图纸不是排版档案！" & vbCrLf & vbCrLf & _
               "请先在 AlphaCAM 中打开排版后的嵌套图纸（ARD 文件，" & vbCrLf & _
               "如 <订单>_<材料>.ard），再点“重新生成标签”。", _
               vbExclamation, "自动化生产排版"
        Exit Sub
    End If
    Set Ni = Nothing

    ' 0.5 未保存修改确认
    '     本流程以【当前内存态】生成标签（先 SaveAs 到临时副本），但结束后会重开
    '     磁盘上的原档案；若用户挪过门板却没保存，就会出现「标签是新的、图纸是
    '     旧的」，未保存的手动移动随之丢失。这里先明确告知，给用户取消去保存的机会。
    If MsgBox("重生成标签将基于【当前图纸的内存状态】。" & vbCrLf & vbCrLf & _
              "如果您刚才手动移动过门板且尚未保存：" & vbCrLf & _
              "  · 新标签会反映这些移动" & vbCrLf & _
              "  · 但重生成结束后重新打开的原档案仍是磁盘上的旧位置，" & vbCrLf & _
              "    未保存的修改将会丢失" & vbCrLf & vbCrLf & _
              "建议先按 Ctrl+S 保存图纸。是否继续？", _
              vbYesNo + vbExclamation, "自动化生产排版") <> vbYes Then
        m_Log "重生成标签：用户取消（未保存修改确认）"
        Exit Sub
    End If

    ' 1. 初始化选项（路径/报表配置与排版时一致）
    Set clsOptions = New COptions
    strCTX = clsOptions.CTXFile

    ' 2. 从刀路属性恢复订单名（m_SetAttributes 写入 DEF_ATT_JOB_NAME）
    gstr_JobName = ""
    For Each P In ActiveDrawing.ToolPaths
        If P.Attribute(DEF_ATT_JOB_NAME) <> "" Then
            gstr_JobName = P.Attribute(DEF_ATT_JOB_NAME)
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
        ' 数据库回退 1：PressDoorImage 路径含真实材料名
        ' 注意：不用 InStrRev 切目录（AlphaCAM VBA 中行为异常），
        ' 直接在全路径中 InStr 定位 JobName 取后续一段
        On Error Resume Next
        If gbln_ConnectToDB() Then
            Set rst = gdb_CDM.Execute("SELECT TOP 1 PressDoorImage FROM AD_REPORT_DATA WHERE PressDoorImage <> '' AND INSTR(PressDoorImage, '" & gs_FixSQL(gstr_JobName) & "') > 0")
            If Not rst Is Nothing Then
                If Not rst.EOF Then
                    sPath = rst.Fields(0)
                    iPos = InStr(sPath, gstr_JobName)
                    If iPos > 0 Then
                        sMat = Mid$(sPath, iPos + Len(gstr_JobName) + 1)
                        If InStr(sMat, "_") > 0 Then sMat = Left$(sMat, InStr(sMat, "_") - 1)
                    End If
                End If
                rst.Close
            End If
        End If
        On Error GoTo EH
    End If
    If sMat = "" Then
        ' 数据库回退 2：明细表 AD_ORDER_DETAILS.Material（单材料订单最可靠）
        On Error Resume Next
        If gbln_ConnectToDB() Then
            Set rst = gdb_CDM.Execute("SELECT DISTINCT Material FROM AD_ORDER_DETAILS WHERE OrderID=(SELECT OrderID FROM AD_ORDERS WHERE JobName='" & gs_FixSQL(gstr_JobName) & "')")
            If Not rst Is Nothing Then
                If Not rst.EOF Then
                    sMat = rst.Fields(0)
                    rst.MoveNext
                    If Not rst.EOF Then sMat = ""   ' 多材料无法确定
                End If
                rst.Close
            End If
        End If
        On Error GoTo EH
    End If
    If sMat = "" Then sMat = ActiveDrawing.Attribute(DEF_ATT_MATERIAL_NAME)
    If sMat = "" Then
        MsgBox "无法恢复材料名（图纸名/路径/明细/属性均未找到）。", _
               vbExclamation, "自动化生产排版"
        Exit Sub
    End If

    ' 4. 数据库连接 + 排版区域集合
    If Not gbln_ConnectToDB() Then
        MsgBox "无法连接 CDM 数据库", vbCritical: Exit Sub
    End If
    m_PopulateNestingZones

    ' 5. 构造材料对象并重新生成 EMF
    Set Material = New CMaterial
    Material.MaterialName = sMat

    ' 5.0 保存用户图到临时副本（含未保存的移动/编辑）
    '     m_CreateAlphaCAMDrawingsOfSheets 在副本上跑，用户原图数据零触碰，
    '     生成后重开洁净副本并写回用户正式文件（保护加工道次关联）
    Set FSO2 = New Scripting.FileSystemObject
    sUserARD = ActiveDrawing.FullName
    sTmpBak = gs_GetCommonAppDataDir & "regen_" & gstr_JobName & "_" & Format$(Timer, "0") & ".ard"
    On Error Resume Next
    ActiveDrawing.SaveAs sTmpBak
    On Error GoTo EH


    ' 5.1 Back up old tag EMFs (move to a backup dir); only removed after a successful regen (rollback-safe)
    sBakDir = gstr_EnsureBackslash(gs_GetCommonAppDataDir) & "regen_backup_" & gstr_JobName & "_" & sMat & "_" & Format$(Timer, "0")
    On Error Resume Next
    MkDir sBakDir
    If Err.Number <> 0 Then sBakDir = ""
    On Error GoTo EH
    If sBakDir <> "" Then
        sBakDir = gstr_EnsureBackslash(sBakDir)
        On Error Resume Next
        sDel = Dir$(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE & DEF_BACKSLASH & _
                     gstr_JobName & DEF_UNDERSCORE & sMat & DEF_UNDERSCORE & "*" & DEF_EXTENSION_EMF)
        Do While sDel <> ""
            Name gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE & DEF_BACKSLASH & sDel As sBakDir & sDel
            sDel = Dir$
        Loop
        On Error GoTo EH
    End If

    ' 5.2 在副本上使用原 CDM 方法生成标签 EMF
    '     m_CreateAlphaCAMDrawingsOfSheets：整板图 + 单板拆分 + 逐件高亮 EMF
    '     （当前件红色+阴影、其余浅灰；只动副本，不动用户原图）
    App.New
    blnScratchOpened = True    ' 用户原图已从窗口撤下：失败时必须帮用户重开
    On Error Resume Next
    App.OpenDrawing sTmpBak
    On Error GoTo EH
    sNestOverride = "regen_scratch_" & gstr_JobName & "_" & Format$(Timer, "0") & DEF_EXTENSION_ARD
    ' 临时巢套 ard 的完整路径在这里就定下来，失败分支(EH)也要用它做清理
    If clsOptions.OutputResultsSubFolder Then
        sScratchFile = gstr_EnsureBackslash(clsOptions.PathToRoot) & gstr_JobName & "\" & sNestOverride
    Else
        sScratchFile = gstr_EnsureBackslash(clsOptions.PathToRoot) & sNestOverride
    End If
    m_CreateAlphaCAMDrawingsOfSheets Material, sNestOverride
    ' 轮询等待 m_Create 真正完成再进入后续。
    ' 期望文件数 = 整板图(每板 1 张) + 逐件图(每板 Parts.Count 张)，见 m_WaitForLabelEMFs 注释
    On Error Resume Next
    lngExpectedDoors = 0
    lngSheetCount = 0
    Set Ni = ActiveDrawing.GetNestInformation
    If Not Ni Is Nothing Then
        For Each Nsh In Ni.Sheets
            lngExpectedDoors = lngExpectedDoors + Nsh.Parts.Count
            lngSheetCount = lngSheetCount + 1
        Next Nsh
    End If
    If Not m_WaitForLabelEMFs(gstr_JobName, sMat, lngExpectedDoors + lngSheetCount) Then
        MsgBox "重生成标签可能未完全完成（等待 EMF 超时）。", vbExclamation, "自动化生产排版"
    End If
    On Error GoTo EH

    ' 5.3 同步 AD_REPORT_DATA：按 DEF_ATT_DETAIL_ID 更新标签路径/件号，删除已删板件记录
    On Error Resume Next
    Set rst2 = gdb_CDM.Execute("SELECT OrderID FROM AD_ORDERS WHERE JobName='" & gs_FixSQL(gstr_JobName) & "'")
    If Not rst2 Is Nothing Then
        If Not rst2.EOF Then lngOrderID2 = rst2.Fields("OrderID")
        rst2.Close
    End If
    If lngOrderID2 > 0 Then
        On Error GoTo EH
        gdb_CDM.BeginTrans
        blnInTrans = True
        sDetailList = ""
        Set Ni = ActiveDrawing.GetNestInformation
        For Each Nsh In Ni.Sheets
            For Each Npi In Nsh.Parts
                For Each P In Npi.Paths
                    If P.Attribute(DEF_ATT_DETAIL_ID) <> "" And P.Attribute(DEF_ATT_NEST_DOOR_COUNT) <> "" Then
                        lngDetail = CLng(P.Attribute(DEF_ATT_DETAIL_ID))
                        lngCnt = CLng(P.Attribute(DEF_ATT_NEST_DOOR_COUNT))
                        sImgPath = gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE & DEF_BACKSLASH & _
                                   gstr_JobName & DEF_UNDERSCORE & sMat & DEF_UNDERSCORE & Nsh.Name & DEF_UNDERSCORE & lngCnt & DEF_EXTENSION_EMF
                        gdb_CDM.Execute "UPDATE AD_REPORT_DATA SET PressDoorImage='" & gs_FixSQL(sImgPath) & "', PressDoorCounter=" & lngCnt & " WHERE DetailID=" & lngDetail
                        If sDetailList = "" Then sDetailList = CStr(lngDetail) Else sDetailList = sDetailList & "," & lngDetail
                    End If
                Next P
            Next Npi
        Next Nsh
        If sDetailList <> "" Then
            gdb_CDM.Execute "DELETE FROM AD_REPORT_DATA WHERE OrderID=" & lngOrderID2 & " AND DetailID NOT IN (" & sDetailList & ")"
        End If
        gdb_CDM.CommitTrans
        blnInTrans = False
        Set Ni = Nothing
    End If
    On Error GoTo EH

    ' 5.3b 成功后删除备份目录（v1.8 修复）
    '      备份目录装的是 5.1 移走的旧 EMF，非空目录 RmDir 必定失败 ——
    '      原实现因此从未真正删成功，ProgramData 累积了 28 个 regen_backup_*
    '      （222 个旧 EMF）。必须先 Kill 再 RmDir。
    If sBakDir <> "" Then
        On Error Resume Next
        Kill sBakDir & "*.*"
        RmDir sBakDir
        On Error GoTo EH
    End If

    ' 5.4 删除临时巢套 ard 与临时副本；用户主图文件从未重存。
    '     真正“切换到原档案”在下面成功弹窗关闭之后执行。
    '     v1.8: 此处的临时嵌套 ard 常仍被 AlphaCAM 占用（m_Create 打开过它），
    '          删除会失败 —— 成功结尾在 App.New 关档之后再补删一次并记日志。
    On Error Resume Next
    If sScratchFile <> "" Then
        If FSO2.FileExists(sScratchFile) Then FSO2.DeleteFile sScratchFile, True
    End If
    If FSO2.FileExists(sTmpBak) Then FSO2.DeleteFile sTmpBak, True
    Set FSO2 = Nothing
    ActiveDrawing.ZoomAll
    On Error GoTo EH

    MsgBox "门板标签 EMF 已重新生成" & vbCrLf & _
           "订单: " & gstr_JobName & vbCrLf & _
           "材料: " & sMat, vbInformation, "自动化生产排版"

    ' 成功弹窗关闭后：先用 App.New 打开空档案（关掉残留 regen 临时窗口），
    ' 再用 App.OpenDrawing（AlphaCAM 打开档案 API）打开用户原档案 <Job>.ard。
    On Error Resume Next
    DoEvents
    Sleep 500
    App.New
    DoEvents
    Sleep 500
    If sUserARD <> "" Then App.OpenDrawing sUserARD
    ActiveDrawing.ZoomAll
    ' v1.8: App.New 已关掉占用临时嵌套 ard 的文档，此时补删 5.4 删不掉的那个文件
    If sScratchFile <> "" Then
        Set FSO2 = New Scripting.FileSystemObject
        If FSO2.FileExists(sScratchFile) Then FSO2.DeleteFile sScratchFile, True
        If FSO2.FileExists(sScratchFile) Then
            m_Log "临时嵌套档案仍未能删除(被占用): " & sScratchFile & " 请手工清理"
        Else
            m_Log "临时嵌套档案已清理: " & sScratchFile
        End If
        Set FSO2 = Nothing
    End If
    ' v1.8: 兜底恢复屏幕/工程栏刷新并重绘（Make.m_CreateAlphaCAMDrawingsOfSheets
    '       会置 False，其恢复语句在 Make.bas:3991/3992 原本被注释掉）
    ActiveDrawing.ScreenUpdating = True
    Frame.ProjectBarUpdating = True
    ActiveDrawing.Redraw
    On Error GoTo EH
    Exit Sub
EH:
    ' ---- 失败现场还原（v1.7）----
    ' 1) 数据库：回滚本次 AD_REPORT_DATA 同步
    If blnInTrans Then
        On Error Resume Next
        gdb_CDM.RollbackTrans
        blnInTrans = False
    End If
    ' 2) 标签文件：把备份目录里的旧 EMF 挪回图片目录
    If sBakDir <> "" Then m_RestoreLabelBackup sBakDir
    ' 3) 临时文件：删除临时副本与临时巢套 ard，别在 ProgramData 留垃圾
    '    （与成功路径同样处理；失败也不影响用户原档案，原档案全程未被写入）
    On Error Resume Next
    If Not FSO2 Is Nothing Then
        If sScratchFile <> "" Then
            If FSO2.FileExists(sScratchFile) Then FSO2.DeleteFile sScratchFile, True
        End If
        If sTmpBak <> "" Then
            If FSO2.FileExists(sTmpBak) Then FSO2.DeleteFile sTmpBak, True
        End If
        Set FSO2 = Nothing
    End If
    ' 4) 现场：只有当 App.New 已经把用户图撤下时才重开原档案，
    '    否则重开会把用户尚未保存的修改一起冲掉
    If blnScratchOpened And sUserARD <> "" Then
        App.New
        DoEvents
        Sleep 500
        App.OpenDrawing sUserARD
        ActiveDrawing.ZoomAll
    End If
    ' v1.8: 失败路径同样兜底恢复刷新，否则报错后画面也是“冻住”的
    ActiveDrawing.ScreenUpdating = True
    Frame.ProjectBarUpdating = True
    ActiveDrawing.Redraw
    On Error GoTo EH
    m_LogError Err.Number, "g_RegenDoorLabelEMFs", Err.Description & " (材料=" & sMat & " 订单=" & gstr_JobName & ")"
    MsgBox "重生成标签失败，已还原现场。" & vbCrLf & vbCrLf & _
           "错误: " & Err.Description & vbCrLf & vbCrLf & _
           "· 数据库同步已回滚" & vbCrLf & _
           "· 旧标签 EMF 已从备份还原" & vbCrLf & _
           IIf(blnScratchOpened And sUserARD <> "", _
               "· 已重新打开原档案" & vbCrLf & "  若您有未保存的手动移动，请检查是否需要重做。", _
               "· 原图纸仍在当前窗口，未做改动"), _
           vbCritical, "自动化生产排版"
End Sub

' ============================================================================
' Restore old tag EMFs from a backup dir back to the image folder (rollback on error)
' ============================================================================
Private Sub m_RestoreLabelBackup(ByVal sBakDir As String)
    Dim sF As String
    Dim sImg As String
    On Error Resume Next
    If sBakDir = "" Then Exit Sub
    sBakDir = gstr_EnsureBackslash(sBakDir)
    sImg = gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE & DEF_BACKSLASH
    sF = Dir$(sBakDir & "*" & DEF_EXTENSION_EMF)
    Do While sF <> ""
        Name sBakDir & sF As sImg & sF
        sF = Dir$
    Loop
    RmDir sBakDir
    On Error GoTo 0
End Sub

