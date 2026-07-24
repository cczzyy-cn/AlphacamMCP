Attribute VB_Name = "BatchProcess"
Option Explicit

' ============================================================================
' BatchProcess — CDM 批量生产模块
' ============================================================================
' 功能: 从 CDM 数据库读取订单门板数据，使用 AlphaCAM API 直接生成几何和刀路
' 依赖: CDM.arb 已加载（提供 gdb_CDM、gbln_ConnectToDB、gs_FixSQL、CNest 等）
' 用法: CDM.BatchProcess.Run 123  (123=OrderID)
'       CDM.BatchProcess.RunByName "订单名称"
' ============================================================================

' ── 入口：按 OrderID 处理 ──
Public Sub Run(ByVal lngOrderID As Long)
    Call ProcessOrder(lngOrderID)
End Sub


' ── 入口：按订单名称查找并处理 ──
Public Sub RunByName(ByVal sJobName As String)
    Dim rst As ADODB.Recordset
    Dim lngOrderID As Long
    
    If Not gbln_ConnectToDB() Then
        MsgBox "无法连接 CDM 数据库", vbCritical
        Exit Sub
    End If
    
    Set rst = gdb_CDM.Execute("SELECT OrderID FROM AD_ORDERS WHERE JobName='" & gs_FixSQL(sJobName) & "'")
    If rst.BOF And rst.EOF Then
        MsgBox "未找到订单: " & sJobName, vbExclamation
        rst.Close
        Exit Sub
    End If
    
    lngOrderID = rst.Fields("OrderID").Value
    rst.Close
    Call ProcessOrder(lngOrderID)
End Sub


' ── 核心处理函数 ──
Private Sub ProcessOrder(ByVal lngOrderID As Long)
    Dim rstDetails As ADODB.Recordset
    Dim sJobName As String
    Dim lngCount As Long
    Dim lngSuccess As Long
    Dim lngFail As Long
    
    On Error GoTo EH
    
    ' 连接数据库
    If Not gbln_ConnectToDB() Then
        MsgBox "无法连接 CDM 数据库", vbCritical
        Exit Sub
    End If
    
    ' 获取订单名
    Dim rstOrder As ADODB.Recordset
    Set rstOrder = gdb_CDM.Execute("SELECT JobName FROM AD_ORDERS WHERE OrderID=" & lngOrderID)
    If rstOrder.BOF And rstOrder.EOF Then
        MsgBox "订单 ID " & lngOrderID & " 不存在", vbExclamation
        rstOrder.Close
        Exit Sub
    End If
    sJobName = rstOrder.Fields("JobName").Value
    rstOrder.Close
    
    ' 读取门板明细
    Set rstDetails = gdb_CDM.Execute( _
        "SELECT * FROM AD_ORDER_DETAILS WHERE OrderID=" & lngOrderID)
    
    If rstDetails.BOF And rstDetails.EOF Then
        MsgBox "订单中没有门板数据", vbExclamation
        rstDetails.Close
        Exit Sub
    End If
    
    rstDetails.MoveLast
    lngCount = rstDetails.RecordCount
    rstDetails.MoveFirst
    
    ' 开始处理
    App.Frame.ShowProgressBox "批量生产: " & sJobName, "准备中..."
    
    lngSuccess = 0
    lngFail = 0
    
    While Not rstDetails.EOF
        Dim dblWidth As Double
        Dim dblLength As Double
        Dim sTypeName As String
        Dim lngQty As Long
        Dim i As Long
        
        sTypeName = gvar_CheckNull(rstDetails.Fields("TypeName"))
        dblWidth = PDbl(gvar_CheckNull(rstDetails.Fields("Width")))
        dblLength = PDbl(gvar_CheckNull(rstDetails.Fields("Length")))
        lngQty = CLng(gvar_CheckNull(rstDetails.Fields("Quantity")))
        If lngQty < 1 Then lngQty = 1
        
        App.Frame.SetProgressText "处理: " & sTypeName & " " & dblWidth & "x" & dblLength
        
        ' 处理每个门板（含数量复制）
        For i = 1 To lngQty
            If Not ProcessDoor(sTypeName, dblWidth, dblLength) Then
                lngFail = lngFail + 1
            Else
                lngSuccess = lngSuccess + 1
            End If
        Next i
        
        rstDetails.MoveNext
    Wend
    
    rstDetails.Close
    
    App.Frame.CloseProgressBox
    
    MsgBox "处理完成!" & vbCrLf & _
           "  成功: " & lngSuccess & vbCrLf & _
           "  失败: " & lngFail, vbInformation, "BatchProcess"
    Exit Sub
    
EH:
    MsgBox "错误: " & Err.Description, vbCritical, "BatchProcess"
    App.Frame.CloseProgressBox
End Sub


' ── 处理单个门板 ──
Private Function ProcessDoor(ByVal sTypeName As String, _
                             ByVal dblWidth As Double, _
                             ByVal dblLength As Double) As Boolean
    Dim pthRect As Path
    Dim rstPaths As ADODB.Recordset
    
    On Error GoTo EH
    
    ' 新建图纸
    App.New
    
    ' 创建门板外框矩形
    Set pthRect = ActiveDrawing.CreateRectangle(0, 0, dblWidth, dblLength)
    pthRect.Attribute("TypeName") = sTypeName
    pthRect.Attribute("Width") = CStr(dblWidth)
    pthRect.Attribute("Length") = CStr(dblLength)
    
    ' 从 AD_DOOR_PATHS 查找该门型的刀路
    Set rstPaths = gdb_CDM.Execute( _
        "SELECT dp.* FROM AD_DOOR_PATHS dp " & _
        "INNER JOIN AD_DOOR_TYPES dt ON dp.TypeID = dt.PK " & _
        "WHERE dt.TypeID='" & gs_FixSQL(sTypeName) & "' " & _
        "ORDER BY dp.PathNumber")
    
    If rstPaths Is Nothing Then
        ' 没有刀路定义，只保存几何
        GoTo SaveDoor
    End If
    
    If rstPaths.BOF And rstPaths.EOF Then
        ' 没有刀路记录
        rstPaths.Close
        GoTo SaveDoor
    End If
    
    ' 遍历刀路并应用
    While Not rstPaths.EOF
        Dim lngProcessType As Long
        Dim dblDepth As Double
        Dim dblStock As Double
        Dim lngToolNum As Long
        Dim dblCutFeed As Double
        Dim dblDownFeed As Double
        Dim lngSpindleSpeed As Long
        Dim dblWidthOfCut As Double
        Dim lngMCComp As Long
        Dim dblSafeRapid As Double
        Dim dblMaterialTop As Double
        Dim dblRapidDownTo As Double
        Dim bAutoZ As Boolean
        Dim dblMaxDepthPerCut As Double
        
        ' 读取刀路参数
        lngProcessType = gvar_CheckNull(rstPaths.Fields("ProcessType"))
        dblDepth = PDbl(gvar_CheckNull(rstPaths.Fields("FinalDepth")))
        dblStock = PDbl(gvar_CheckNull(rstPaths.Fields("StockAllowance")))
        lngToolNum = CLng(gvar_CheckNull(rstPaths.Fields("ToolNumber")))
        dblCutFeed = PDbl(gvar_CheckNull(rstPaths.Fields("CutFeed")))
        dblDownFeed = PDbl(gvar_CheckNull(rstPaths.Fields("DownFeed")))
        lngSpindleSpeed = CLng(gvar_CheckNull(rstPaths.Fields("SpindleSpeed")))
        dblWidthOfCut = PDbl(gvar_CheckNull(rstPaths.Fields("WidthOfCut")))
        lngMCComp = CInt(gvar_CheckNull(rstPaths.Fields("MCComp")))
        dblSafeRapid = PDbl(gvar_CheckNull(rstPaths.Fields("SafeRapidLevel")))
        dblMaterialTop = PDbl(gvar_CheckNull(rstPaths.Fields("MaterialTop")))
        dblRapidDownTo = PDbl(gvar_CheckNull(rstPaths.Fields("RapidDownTo")))
        bAutoZ = CBool(gvar_CheckNull(rstPaths.Fields("AutoZ")))
        dblMaxDepthPerCut = PDbl(gvar_CheckNull(rstPaths.Fields("MaxDepthPerCut")))
        
        ' 选择刀具
        Dim lngToolID As Long
        lngToolID = CLng(gvar_CheckNull(rstPaths.Fields("ToolID")))
        If lngToolID > 0 Then
            ' 使用 CDM 的刀具选择逻辑
            SelectToolByID lngToolID
        ElseIf lngToolNum > 0 Then
            App.SelectTool CStr(lngToolNum)
        End If
        
        ' 应用刀路到几何
        pthRect.Selected = True
        
        Dim Md As MillData
        Set Md = App.ActiveDrawing.CreateMillData
        With Md
            .ProcessType2 = lngProcessType
            .FinalDepth = dblDepth
            .StockAllowance = dblStock
            .CutFeed = dblCutFeed
            .DownFeed = dblDownFeed
            .SpindleSpeed = lngSpindleSpeed
            .WidthOfCut = dblWidthOfCut
            .MCComp = lngMCComp
            .SafeRapidLevel = dblSafeRapid
            .MaterialTop = dblMaterialTop
            .RapidDownTo = dblRapidDownTo
            If dblMaxDepthPerCut > 0 Then
                .MaxDepthPerCut = dblMaxDepthPerCut
            End If
        End Set
        
        ' 执行加工
        App.ActiveDrawing.ApplyMillData Md
        
        rstPaths.MoveNext
    Wend
    
    rstPaths.Close
    
SaveDoor:
    ' 保存图纸
    Dim sFileName As String
    sFileName = App.Frame.PathOfThisAddin & "\" & sTypeName & "_" & _
                Replace(CStr(dblWidth), ".", "_") & "x" & _
                Replace(CStr(dblLength), ".", "_") & ".amd"
    
    ActiveDrawing.SaveAs sFileName
    ProcessDoor = True
    Exit Function
    
EH:
    ProcessDoor = False
End Function


' ── 根据 ToolID 选择刀具 ──
Private Sub SelectToolByID(ByVal lngToolID As Long)
    Dim rstTool As ADODB.Recordset
    
    On Error Resume Next
    
    Set rstTool = gdb_CDM.Execute("SELECT * FROM AD_TOOLS WHERE ToolPK=" & lngToolID)
    If rstTool Is Nothing Then Exit Sub
    If rstTool.BOF And rstTool.EOF Then
        rstTool.Close
        Exit Sub
    End If
    
    Dim sToolName As String
    sToolName = gvar_CheckNull(rstTool.Fields("ToolName"))
    rstTool.Close
    
    If sToolName <> "" Then
        App.SelectTool sToolName
    End If
End Sub
