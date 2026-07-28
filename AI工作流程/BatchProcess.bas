Attribute VB_Name = "BatchProcess"
Option Explicit

' ============================================================================
' BatchProcess — CDM 批量生产模块
' ============================================================================
' 功能: 从 CDM 数据库读取订单门板数据，生成几何图形并保存为 .amd 文件
' 依赖: CDM.arb 已加载（提供 gdb_CDM、gbln_ConnectToDB、gs_FixSQL 等）
' 用法: CDM.BatchProcess.Run 123           (123=OrderID)
'       CDM.BatchProcess.RunByName "订单名" (按名称查找)
' ============================================================================

Public Sub Run(ByVal lngOrderID As Long)
    Call ProcessOrder(lngOrderID)
End Sub

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

Private Sub ProcessOrder(ByVal lngOrderID As Long)
    Dim rstDetails As ADODB.Recordset
    Dim sJobName As String
    Dim lngSuccess As Long, lngFail As Long
    On Error GoTo EH
    If Not gbln_ConnectToDB() Then
        MsgBox "无法连接 CDM 数据库", vbCritical
        Exit Sub
    End If
    Dim rstOrder As ADODB.Recordset
    Set rstOrder = gdb_CDM.Execute("SELECT JobName FROM AD_ORDERS WHERE OrderID=" & lngOrderID)
    If rstOrder.BOF And rstOrder.EOF Then
        MsgBox "订单 ID " & lngOrderID & " 不存在", vbExclamation
        rstOrder.Close
        Exit Sub
    End If
    sJobName = rstOrder.Fields("JobName").Value
    rstOrder.Close
    Set rstDetails = gdb_CDM.Execute("SELECT * FROM AD_ORDER_DETAILS WHERE OrderID=" & lngOrderID)
    If rstDetails.BOF And rstDetails.EOF Then
        MsgBox "订单中没有门板数据", vbExclamation
        rstDetails.Close
        Exit Sub
    End If
    rstDetails.MoveLast
    Dim lngCount As Long: lngCount = rstDetails.RecordCount
    rstDetails.MoveFirst
    App.Frame.ShowProgressBox "批量生产: " & sJobName, "准备中..."
    lngSuccess = 0: lngFail = 0
    While Not rstDetails.EOF
        Dim sTypeName As String, dblWidth As Double, dblLength As Double, lngQty As Long, i As Long
        sTypeName = gvar_CheckNull(rstDetails.Fields("TypeName"))
        dblWidth = PDbl(gvar_CheckNull(rstDetails.Fields("Width")))
        dblLength = PDbl(gvar_CheckNull(rstDetails.Fields("Length")))
        lngQty = CLng(gvar_CheckNull(rstDetails.Fields("Quantity")))
        If lngQty < 1 Then lngQty = 1
        App.Frame.SetProgressText "处理: " & sTypeName & " " & dblWidth & "x" & dblLength
        For i = 1 To lngQty
            If ProcessDoor(sTypeName, dblWidth, dblLength) Then
                lngSuccess = lngSuccess + 1
            Else
                lngFail = lngFail + 1
            End If
        Next i
        rstDetails.MoveNext
    Wend
    rstDetails.Close
    App.Frame.CloseProgressBox
    MsgBox "处理完成!" & vbCrLf & "  成功: " & lngSuccess & vbCrLf & "  失败: " & lngFail, vbInformation, "BatchProcess"
    Exit Sub
EH:
    MsgBox "错误: " & Err.Description, vbCritical, "BatchProcess"
    App.Frame.CloseProgressBox
End Sub

Private Function ProcessDoor(ByVal sTypeName As String, ByVal dblWidth As Double, ByVal dblLength As Double) As Boolean
    Dim pthRect As Path
    On Error GoTo EH
    App.New
    Set pthRect = ActiveDrawing.CreateRectangle(0, 0, dblWidth, dblLength)
    pthRect.Attribute("TypeName") = sTypeName
    pthRect.Attribute("Width") = CStr(dblWidth)
    pthRect.Attribute("Length") = CStr(dblLength)
    Dim sFileName As String
    sFileName = App.Frame.PathOfThisAddin & "\" & sTypeName & "_" & Replace(CStr(dblWidth), ".", "_") & "x" & Replace(CStr(dblLength), ".", "_") & ".amd"
    ActiveDrawing.SaveAs sFileName
    ProcessDoor = True
    Exit Function
EH:
    ProcessDoor = False
End Function
