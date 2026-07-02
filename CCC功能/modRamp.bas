' ==============================================================================
' CCC功能 — modRamp 斜角下刀
' ==============================================================================
' 算法参照"小条先切"：
'   1. 刀具路径包围盒判断小板件（单边 < 范围）
'   2. SetLeadInOutAuto(SlopeIn=True) 实现斜坡入刀
'   3. 深度 ≥ 20mm 时先粗切 60% 再精切（参照小条先切）
' ==============================================================================
Option Explicit
Option Private Module

Private Const ATT_RAMP_DONE    As String = "CCC_RampDone"
Private Const DEG2RAD          As Double = 0.0174532925199433

' ==============================================================================
' Sub 斜角下刀() — 入口（由 Events.bas 菜单调用）
' ==============================================================================
Sub 斜角下刀()
    frmRamp.Show vbModeless
End Sub

' ==============================================================================
' Public Sub ApplyRampEntry() — 执行斜角下刀算法
' ==============================================================================
Public Sub ApplyRampEntry(ByVal minSize As Double, _
                          ByVal cutDepth As Double, _
                          ByVal rampAngle As Double, _
                          ByVal methodName As String, _
                          ByVal toolMatch As String, _
                          Optional ByVal tNum As Long = 0)

    On Error GoTo ErrHandler

    Dim drw As Drawing
    Dim ni As NestInformation
    Dim ops As Operations
    Dim i As Long, j As Long, k As Long
    Dim op As Operation
    Dim subs As SubOperations
    Dim subop As SubOperation
    Dim mt As MillTool
    Dim tps As Paths
    Dim tp As Path
    Dim totalCount As Long, smallCount As Long, rampApplied As Long
    Dim skipCount As Long
    Dim tpW As Double, tpH As Double
    Dim toolRadius As Double
    Dim rampLength As Double, lengthMult As Double
    Dim isMatch As Boolean
    Dim spPos As Integer
    Dim procName As String
    Dim selToolNum As Long

    Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        MsgBox "没有活动图纸！", vbExclamation, "斜角下刀"
        Exit Sub
    End If

    drw.ScreenUpdating = False
    App.SetUndoCommandName "斜角下刀"
    App.SetUndoPoint

    Set ni = drw.GetNestInformation
    Set ops = drw.Operations
    If ops Is Nothing Or ops.Count = 0 Then
        drw.ScreenUpdating = True
        MsgBox "图纸中没有加工操作！", vbExclamation, "斜角下刀"
        Exit Sub
    End If

    totalCount = 0: smallCount = 0: rampApplied = 0: skipCount = 0

    For i = 1 To ops.Count
        Set op = ops(i)
        Set subs = op.SubOperations
        If subs Is Nothing Then GoTo NextOp

        For j = 1 To subs.Count
            Set subop = subs(j)
            Set mt = subop.Tool
            If mt Is Nothing Then GoTo NextSub

            ' 提取加工方式名
            procName = subop.Name
            spPos = InStr(procName, "  ")
            If spPos > 0 Then procName = Left(procName, spPos - 1) _
            Else: spPos = InStr(procName, " "): If spPos > 0 Then procName = Left(procName, spPos - 1)
            If methodName <> "" And procName <> methodName Then GoTo NextSub

            ' 刀具匹配
            isMatch = False
            If toolMatch <> "" Then
                If mt.Name = toolMatch Then isMatch = True
                ElseIf InStr(1, mt.Name, toolMatch, vbTextCompare) > 0 Then isMatch = True
                ElseIf InStr(1, toolMatch, mt.Name, vbTextCompare) > 0 Then isMatch = True
                ElseIf CStr(mt.Number) = toolMatch Then isMatch = True
                ElseIf tNum > 0 And mt.Number = tNum Then isMatch = True
                ElseIf Left(toolMatch, 1) = "T" Then
                    selToolNum = Val(Mid(toolMatch, 2))
                    If selToolNum > 0 And mt.Number = selToolNum Then isMatch = True
                End If
            Else
                isMatch = True
            End If
            If Not isMatch Then GoTo NextSub

            Set tps = subop.ToolPaths
            If tps Is Nothing Then GoTo NextSub

            For k = 1 To tps.Count
                Set tp = tps(k)
                If tp Is Nothing Then GoTo NextTp
                If tp.Attribute(ATT_RAMP_DONE) <> 0 Then
                    skipCount = skipCount + 1: GoTo NextTp
                End If

                totalCount = totalCount + 1

                ' === 刀具路径包围盒判断小板件 ===
                tpW = tp.MaxXL - tp.MinXL
                tpH = tp.MaxYL - tp.MinYL

                If tpW > 1 And tpH > 1 And (tpW < minSize Or tpH < minSize) Then
                    smallCount = smallCount + 1

                    toolRadius = 3
                    If Not (mt Is Nothing) Then
                        toolRadius = mt.Diameter / 2
                        If toolRadius <= 0 Then toolRadius = 3
                    End If

                    ' 斜坡长度 = 切割深度 / Tan(角度)
                    rampLength = cutDepth / Tan(rampAngle * DEG2RAD)
                    lengthMult = rampLength / toolRadius
                    If lengthMult < 1 Then lengthMult = 1
                    If lengthMult > 50 Then lengthMult = 50

                    ' === 深度 ≥ 20mm 时先粗切 60% 深度（参照小条先切） ===
                    If cutDepth >= 20 Then
                        Call RoughPass60Percent subop, cutDepth, toolRadius
                    End If

                    ' === 应用斜角下刀 ===
                    ' SetLeadInOutAuto: SlopeIn=True 使刀具沿入刀线斜坡下降
                    ' Overlap=-0.5 留 0.5mm 连接点
                    tp.SetLeadInOutAuto acamLeadBOTH, acamLeadNONE, _
                                        lengthMult, 0.5, 45, _
                                        True, False, -0.5

                    tp.Attribute(ATT_RAMP_DONE) = 1
                    rampApplied = rampApplied + 1
                End If

NextTp:
            Next k
NextSub:
        Next j
NextOp:
    Next i

    drw.ScreenUpdating = True
    drw.Redraw
    If rampApplied > 0 Then drw.ZoomAll: DoEvents

    Dim msg As String
    msg = "斜角下刀处理完成！" & vbCrLf & vbCrLf & _
          "匹配刀具路径: " & totalCount & " 条" & vbCrLf & _
          "其中小板件: " & smallCount & " 条" & vbCrLf & _
          "已应用斜角下刀: " & rampApplied & " 条" & vbCrLf & _
          "（跳过已处理: " & skipCount & " 条）" & vbCrLf & vbCrLf & _
          "参数: 小条范围=" & minSize & "mm" & vbCrLf & _
          "      切割深度=" & cutDepth & "mm" & vbCrLf & _
          "      下刀角度=" & rampAngle & ChrW(176)
    If methodName <> "" Then msg = msg & vbCrLf & "      加工方式=" & methodName
    MsgBox msg, vbInformation, "斜角下刀"
    Exit Sub

ErrHandler:
    If Not (drw Is Nothing) Then drw.ScreenUpdating = True: drw.Redraw
    MsgBox "斜角下刀出错：" & Err.Description, vbCritical, "斜角下刀"
End Sub

' ==============================================================================
' RoughPass60Percent — 深板（≥20mm）先行粗切 60% 深度
' 参照小条先切 CutToolPath.RoughFinish 中 FinalDepth <= -20 的处理
' ==============================================================================
Private Sub RoughPass60Percent(ByVal subop As SubOperation, _
                               ByVal cutDepth As Double, _
                               ByVal toolRadius As Double)

    On Error Resume Next

    Dim geos As Paths: Set geos = subop.Geometries
    If geos Is Nothing Then Exit Sub
    If geos.Count = 0 Then Exit Sub

    Dim geo As Path: Set geo = geos(1)
    If geo Is Nothing Then Exit Sub

    Dim md As MillData: Set md = subop.GetMillData
    If md Is Nothing Then Exit Sub

    Dim mdRough As MillData
    Set mdRough = App.CreateMillData
    mdRough.SafeRapidLevel = md.SafeRapidLevel
    mdRough.RapidDownTo = md.RapidDownTo
    mdRough.SpindleSpeed = 24000
    mdRough.CutFeed = 9000
    mdRough.DownFeed = 2000
    mdRough.FinalDepth = -cutDepth * 0.6

    geo.Selected = True
    mdRough.RoughFinish
    geo.Selected = False

    Set mdRough = Nothing
End Sub
