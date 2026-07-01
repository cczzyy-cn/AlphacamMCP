' ==============================================================================
' CCC功能 — modRamp 斜角下刀
' ==============================================================================
' 功能：对小门板或窄条（单边 < 指定范围）的轮廓刀具路径，
'       应用斜坡入刀（SetLeadInOutAuto with SlopeIn），
'       使刀具沿倾斜路径切入，保留连接处逐渐减少最后吃刀量。
' ==============================================================================
Option Explicit
Option Private Module

' --- 属性常量 ---
Private Const ATT_RAMP_DONE          As String = "CCC_RampDone"
Private Const ATT_SHEET_IDENT        As String = "LicomUKsab_sheet_ident"
Private Const DEG2RAD                As Double = 0.0174532925199433   ' PI/180

' ==============================================================================
' Sub 斜角下刀() — 入口（由 Events.bas 菜单调用）
' ==============================================================================
Sub 斜角下刀()
    frmRamp.Show vbModeless
End Sub


' ==============================================================================
' Public Sub ApplyRampEntry() — 执行斜角下刀算法
'
' 参数：
'   minSize    — 小条范围（单边小于此值视为小板件）
'   cutDepth   — 切割深度（板材厚度）
'   rampAngle  — 下刀角度（度）
'   methodName — 加工方式名称（如 "粗加工" / "精加工"，空=不限）
'   toolMatch  — 刀具匹配串（"T# 刀具名"）
' ==============================================================================
Public Sub ApplyRampEntry(ByVal minSize As Double, _
                          ByVal cutDepth As Double, _
                          ByVal rampAngle As Double, _
                          ByVal methodName As String, _
                          ByVal toolMatch As String, _
                          Optional ByVal tNum As Long = 0)

    On Error GoTo ErrHandler

    ' --- 集中声明 ---
    Dim drw As Drawing
    Dim ni As NestInformation
    Dim sh As NestSheet
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
    Dim tpMinX As Double, tpMaxX As Double
    Dim tpMinY As Double, tpMaxY As Double
    Dim tpW As Double, tpH As Double
    Dim sheetCX As Double, sheetCY As Double
    Dim pcx As Double, pcy As Double
    Dim dx As Double, dy As Double
    Dim sideName As String
    Dim toolRadius As Double
    Dim rampLength As Double
    Dim lengthMult As Double
    Dim radiusMult As Double
    Dim tpName As String
    Dim isMatch As Boolean
    Dim spPos As Integer
    Dim thisMethod As String
    Dim thisToolName As String
    Dim thisToolNum As Long
    Dim selToolNum As Long
    Dim procName As String
    Dim foundSheet As Boolean

    ' --- 基本检查 ---
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then
        MsgBox "没有活动图纸！", vbExclamation, "斜角下刀"
        Exit Sub
    End If

    ' 锁定屏幕
    drw.ScreenUpdating = False
    App.SetUndoCommandName "斜角下刀"
    App.SetUndoPoint

    ' --- 获取 NestInformation 和排版中心 ---
    Set ni = drw.GetNestInformation
    If ni Is Nothing Then
        MsgBox "当前图纸没有排版信息，无法确定排版中心。" & vbCrLf & _
               "将使用图纸全局边界作为排版区域。", vbInformation, "斜角下刀"
    End If

    totalCount = 0
    smallCount = 0
    rampApplied = 0
    skipCount = 0

    ' --- 遍历 Operations → SubOperations → ToolPaths ---
    Set ops = drw.Operations

    If ops Is Nothing Or ops.Count = 0 Then
        ' 无 Operations，直接遍历所有刀具路径
        Dim tpIdx As Long
        Dim tpCnt As Long: tpCnt = drw.GetToolPathCount
        If tpCnt = 0 Then
            drw.ScreenUpdating = True
            MsgBox "图纸中没有刀具路径！", vbExclamation, "斜角下刀"
            Exit Sub
        End If

        Dim tpItem As Path: Set tpItem = drw.GetFirstToolPath
        For tpIdx = 1 To tpCnt
            If tpItem Is Nothing Then GoTo NextTpDirect

            ' 跳过已处理的路径
            If tpItem.Attribute(ATT_RAMP_DONE) <> 0 Then
                skipCount = skipCount + 1
                GoTo NextTpDirect
            End If

            ' 刀具匹配
            Set mt = tpItem.GetTool
            isMatch = False
            If Not (mt Is Nothing) Then
                If toolMatch <> "" Then
                    ' 精确匹配 → 包含匹配（双向）→ T号匹配
                    If mt.Name = toolMatch Then
                        isMatch = True
                    ElseIf InStr(1, mt.Name, toolMatch, vbTextCompare) > 0 Then
                        isMatch = True
                    ElseIf InStr(1, toolMatch, mt.Name, vbTextCompare) > 0 Then
                        isMatch = True
                    ElseIf CStr(mt.Number) = toolMatch Then
                        isMatch = True
                    ElseIf tNum > 0 And mt.Number = tNum Then
                        isMatch = True
                    ElseIf Left(toolMatch, 1) = "T" Then
                        selToolNum = Val(Mid(toolMatch, 2))
                        If selToolNum > 0 And mt.Number = selToolNum Then isMatch = True
                    End If
                Else
                    isMatch = True    ' 不限刀具
                End If
            End If

            If isMatch Then
                totalCount = totalCount + 1

                ' 判断是否为小板件
                tpMinX = tpItem.MinXL: tpMaxX = tpItem.MaxXL
                tpMinY = tpItem.MinYL: tpMaxY = tpItem.MaxYL
                tpW = tpMaxX - tpMinX
                tpH = tpMaxY - tpMinY

                If tpW < minSize Or tpH < minSize Then
                    smallCount = smallCount + 1

                    ' 计算刀具半径
                    toolRadius = 3    ' 默认 3mm
                    Set mt = tpItem.GetTool
                    If Not (mt Is Nothing) Then
                        toolRadius = mt.Diameter / 2
                        If toolRadius <= 0 Then toolRadius = 3
                    End If

                    ' 计算斜坡长度: rampLength = cutDepth / Tan(rampAngle)
                    rampLength = cutDepth / Tan(rampAngle * DEG2RAD)
                    lengthMult = rampLength / toolRadius
                    If lengthMult < 1 Then lengthMult = 1
                    If lengthMult > 50 Then lengthMult = 50

                    radiusMult = 0.5

                    ' 应用斜坡入刀
                    tpItem.SetLeadInOutAuto acamLeadBOTH, acamLeadNONE, _
                                            lengthMult, radiusMult, 45, _
                                            True, False, -0.5

                    ' 标记已处理
                    tpItem.Attribute(ATT_RAMP_DONE) = 1

                    rampApplied = rampApplied + 1
                End If
            End If

NextTpDirect:
            Set tpItem = tpItem.GetNext
        Next tpIdx

        GoTo Report
    End If

    ' --- 通过 Operations 遍历 ---
    For i = 1 To ops.Count
        Set op = ops(i)
        Set subs = op.SubOperations
        If subs Is Nothing Then GoTo NextOp

        For j = 1 To subs.Count
            Set subop = subs(j)
            Set mt = subop.Tool
            If mt Is Nothing Then GoTo NextSub

            ' 提取加工方式名称
            procName = subop.Name
            spPos = InStr(procName, "  ")
            If spPos > 0 Then
                procName = Left(procName, spPos - 1)
            Else
                spPos = InStr(procName, " ")
                If spPos > 0 Then procName = Left(procName, spPos - 1)
            End If

            ' 加工方式匹配
            If methodName <> "" And procName <> methodName Then GoTo NextSub

            ' 刀具匹配
            isMatch = False
            If toolMatch <> "" Then
                ' 精确匹配 → 包含匹配（双向）→ T号匹配
                If mt.Name = toolMatch Then
                    isMatch = True
                ElseIf InStr(1, mt.Name, toolMatch, vbTextCompare) > 0 Then
                    isMatch = True
                ElseIf InStr(1, toolMatch, mt.Name, vbTextCompare) > 0 Then
                    isMatch = True
                ElseIf CStr(mt.Number) = toolMatch Then
                    isMatch = True
                ElseIf tNum > 0 And mt.Number = tNum Then
                    isMatch = True
                ElseIf Left(toolMatch, 1) = "T" Then
                    selToolNum = Val(Mid(toolMatch, 2))
                    If selToolNum > 0 And mt.Number = selToolNum Then isMatch = True
                End If
            Else
                isMatch = True    ' 不限刀具
            End If

            If Not isMatch Then GoTo NextSub

            ' 遍历此 SubOperation 的所有刀具路径
            Set tps = subop.ToolPaths
            If tps Is Nothing Then GoTo NextSub

            For k = 1 To tps.Count
                Set tp = tps(k)
                If tp Is Nothing Then GoTo NextTp

                ' 跳过已处理的路径
                If tp.Attribute(ATT_RAMP_DONE) <> 0 Then
                    skipCount = skipCount + 1
                    GoTo NextTp
                End If

                totalCount = totalCount + 1

                ' 判断是否为小板件：用刀具路径的包围盒
                tpMinX = tp.MinXL: tpMaxX = tp.MaxXL
                tpMinY = tp.MinYL: tpMaxY = tp.MaxYL
                tpW = tpMaxX - tpMinX
                tpH = tpMaxY - tpMinY

                If tpW < minSize Or tpH < minSize Then
                    smallCount = smallCount + 1

                    ' 计算刀具半径
                    toolRadius = 3
                    If Not (mt Is Nothing) Then
                        toolRadius = mt.Diameter / 2
                        If toolRadius <= 0 Then toolRadius = 3
                    End If

                    ' 计算斜坡长度: rampLength = cutDepth / Tan(rampAngle)
                    rampLength = cutDepth / Tan(rampAngle * DEG2RAD)
                    lengthMult = rampLength / toolRadius
                    If lengthMult < 1 Then lengthMult = 1
                    If lengthMult > 50 Then lengthMult = 50
                    radiusMult = 0.5

                    ' 获取排版信息：确定排版中心
                    sheetCX = 0: sheetCY = 0
                    foundSheet = False
                    If Not (ni Is Nothing) Then
                        Dim geo As Path
                        For Each sh In ni.Sheets
                            ' 检查此刀路是否属于本 Sheet
                            Dim pathsInSheet As Paths
                            Set pathsInSheet = sh.Paths
                            If Not (pathsInSheet Is Nothing) Then
                                Dim pi As Long
                                For pi = 1 To pathsInSheet.Count
                                    If pathsInSheet(pi).OpNo = tp.OpNo Then
                                        ' 找到所属 Sheet
                                        Set geo = sh.Geometry
                                        If Not (geo Is Nothing) Then
                                            sheetCX = (geo.MinXL + geo.MaxXL) / 2
                                            sheetCY = (geo.MinYL + geo.MaxYL) / 2
                                        End If
                                        foundSheet = True
                                        Exit For
                                    End If
                                Next pi
                            End If
                            If foundSheet Then Exit For
                        Next sh
                    End If

                    ' 未找到排版中心时的兜底：遍历所有几何求全局边界
                    If Not foundSheet Then
                        Dim geoIdx As Long
                        Dim geoP As Path
                        Dim gMinX As Double, gMaxX As Double
                        Dim gMinY As Double, gMaxY As Double
                        gMinX = 1E+20: gMaxX = -1E+20
                        gMinY = 1E+20: gMaxY = -1E+20
                        Set geoP = drw.GetFirstGeo
                        For geoIdx = 1 To drw.GetGeoCount
                            If Not (geoP Is Nothing) Then
                                If geoP.MinXL < gMinX Then gMinX = geoP.MinXL
                                If geoP.MaxXL > gMaxX Then gMaxX = geoP.MaxXL
                                If geoP.MinYL < gMinY Then gMinY = geoP.MinYL
                                If geoP.MaxYL > gMaxY Then gMaxY = geoP.MaxYL
                                Set geoP = geoP.GetNext
                            End If
                        Next geoIdx
                        If gMaxX > gMinX And gMaxY > gMinY Then
                            sheetCX = (gMinX + gMaxX) / 2
                            sheetCY = (gMinY + gMaxY) / 2
                        Else
                            sheetCX = 0: sheetCY = 0
                        End If
                    End If

                    ' 确定朝向排版中心的那边（用于提示信息）
                    pcx = (tpMinX + tpMaxX) / 2
                    pcy = (tpMinY + tpMaxY) / 2
                    dx = sheetCX - pcx
                    dy = sheetCY - pcy

                    sideName = "右"
                    If Abs(dx) > Abs(dy) Then
                        If dx > 0 Then sideName = "右" Else sideName = "左"
                    Else
                        If dy > 0 Then sideName = "上" Else sideName = "下"
                    End If

                    ' ==========================================================
                    ' 应用斜角下刀：斜坡入刀 + 留连接点
                    ' ==========================================================
                    ' LeadIn 使用 acamLeadBOTH（直线+圆弧）以实现平滑入刀
                    ' SlopeIn=True 使刀具沿入刀线斜坡下降
                    ' Overlap=-0.5 留 0.5mm 负重叠（连接点/tag）
                    tp.SetLeadInOutAuto acamLeadBOTH, acamLeadNONE, _
                                        lengthMult, radiusMult, 45, _
                                        True, False, -0.5

                    ' 标记已处理
                    tp.Attribute(ATT_RAMP_DONE) = 1

                    rampApplied = rampApplied + 1
                End If

NextTp:
            Next k
NextSub:
        Next j
NextOp:
    Next i

Report:
    ' 恢复屏幕更新
    drw.ScreenUpdating = True
    drw.Redraw
    If rampApplied > 0 Then
        drw.ZoomAll
        drw.Refresh
        DoEvents
    End If

    ' 显示结果
    Dim msg As String
    msg = "斜角下刀处理完成！" & vbCrLf & vbCrLf & _
          "匹配刀具路径: " & totalCount & " 条" & vbCrLf & _
          "其中小板件: " & smallCount & " 条" & vbCrLf & _
          "已应用斜角下刀: " & rampApplied & " 条" & vbCrLf & _
          "（跳过已处理: " & skipCount & " 条）" & vbCrLf & vbCrLf & _
          "参数: 小条范围=" & minSize & "mm" & vbCrLf & _
          "      切割深度=" & cutDepth & "mm" & vbCrLf & _
          "      下刀角度=" & rampAngle & "路"

    ' 如已勾选加工方式，追加提示信息
    If methodName <> "" Then
        msg = msg & vbCrLf & "      加工方式=" & methodName
    End If

    MsgBox msg, vbInformation, "斜角下刀"

    Exit Sub

ErrHandler:
    If Not (drw Is Nothing) Then
        drw.ScreenUpdating = True
        drw.Redraw
    End If
    MsgBox "斜角下刀出错：" & Err.Description & vbCrLf & _
           "错误代码: " & Err.Number, vbCritical, "斜角下刀"
End Sub
