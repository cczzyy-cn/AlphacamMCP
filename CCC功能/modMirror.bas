' ==============================================================================
' CCC功能 — modMirror 反面镜像
' ==============================================================================
' 算法：
'   镜像 BM 刀具路径到反面位置，删除正面版件原路径
'   镜像路径的 Part Note 属性设为 "rev"，方便在加工道次中识别
' ==============================================================================
Option Explicit
Option Private Module

' --- Nesting 属性常量（与 RevNest 一致）---
Private Const ATT_PATH_FILE          As String = "LicomUKsab_nest_path_file"
Private Const ATT_FIRST_PATH         As String = "LicomUKsab_nest_first_path"
Private Const ATT_REQUIRED           As String = "LicomUKsab_nest_required"
Private Const ATT_SHEET_IDENT        As String = "LicomUKsab_sheet_ident"
Private Const ATT_SHEET_MATERIAL     As String = "LicomUKsab_sheet_material"
Private Const ATT_SHEET_THICKNESS    As String = "LicomUKsab_sheet_thickness"
Private Const ATT_PART_MOVEX         As String = "LicomUKsab_part_movex"
Private Const ATT_PART_MOVEY         As String = "LicomUKsab_part_movey"
Private Const ATT_PART_ROTANGLE      As String = "LicomUKsab_part_rotangle"
Private Const ATT_PART_MIRRORED      As String = "LicomUKsab_part_mirrored"
Private Const ATT_IS_BOBBLE          As String = "LicomUKsab_is_bobble"
Private Const ATT_PART_MOVE_BY_X     As String = "LicomUKja_part_move_by_x"
Private Const ATT_PART_MOVE_BY_Y     As String = "LicomUKja_part_move_by_y"
Private Const ATT_PART_SHIFT_X       As String = "LicomUKja_part_shift_x"
Private Const ATT_PART_SHIFT_Y       As String = "LicomUKja_part_shift_y"
Private Const ATT_NEST_ITEM_NUM      As String = "LicomUKsab_nest_item_number"
Private Const ATT_IS_REV_SIDE        As String = "AcamUSrg_IsReverseSide"
Private Const ATT_REV_TEXT           As String = "AcamUSrg_TextIsReversed"
Private Const ATT_PART_NOTE          As String = "LicomJPhtOpListPartsName"


' ==============================================================================
Sub 反面镜像()
    fMirror.Show vbModeless
End Sub


' ==============================================================================
' DoMirror — 执行反面镜像（由 fMirror.cmdOK_Click 调用）
' mirrorX:       True=绕X轴（水平线，垂直翻转），False=绕Y轴（垂直线，水平翻转）
' mirrorID:      刀具名过滤标识（空=不过滤全部镜像，非空=仅匹配刀具名的路径）
' ==============================================================================
Public Sub DoMirror(ByVal mirrorX As Boolean, _
                    Optional ByVal mirrorID As String = "")
    
    On Error Resume Next
    
    Dim Drw As Drawing
    Dim P As Path
    Dim pcopy As Path
    Dim minx As Double, maxx As Double
    Dim miny As Double, maxy As Double
    Dim mirrorVal As Double
    Dim tp As Path
    Dim ni As NestInformation
    Dim sh As NestSheet
    Dim I As Long
    Dim lastop As Long
    Dim maxop As Long, minop As Long
    Dim mTool As MillTool
    Dim tpIdx As Long, tpCount As Long, mirroredCount As Long
    Dim collectTP() As Path
    
    Set Drw = App.ActiveDrawing
    Set ni = Drw.GetNestInformation
    
    ' --- 基本检查 ---
    If ni.Sheets.count = 0 Then
        MsgBox "当前图纸没有排版信息，无法执行反面镜像。", vbInformation, "反面镜像"
        GoTo byebye
    End If
    If Drw.GetToolPathCount = 0 Then
        MsgBox "没有找到刀具路径。", vbInformation, "反面镜像"
        GoTo byebye
    End If
    
    ' 锁定屏幕更新
    Drw.ScreenUpdating = False
    App.SetUndoCommandName "反面镜像"
    App.SetUndoPoint
    
    ' --- 计算 Sheet 包围盒 ---
    minx = 1E+20: miny = 1E+20
    maxx = -1E+20: maxy = -1E+20
    For Each sh In ni.Sheets
        Set P = sh.Geometry
        If P.MinXL < minx Then minx = P.MinXL
        If P.MinYL < miny Then miny = P.MinYL
        If P.MaxXL > maxx Then maxx = P.MaxXL
        If P.MaxYL > maxy Then maxy = P.MaxYL
    Next sh
    
    ' --- 确定镜像线位置（RevNest 方式：偏移 5% 边界外）---
    If mirrorX Then
        mirrorVal = minx - ((maxx - minx) * 0.05)
    Else
        mirrorVal = miny - ((maxy - miny) * 0.05)
    End If
    
    ' ======================================================================
    ' Phase 2 — 镜像 BM 刀具路径，删除正面版件原路径
    ' ======================================================================
    ' 收集主图纸中匹配的刀具路径（因为后面要删除，先收集再处理）
    tpCount = 0
    
    Set tp = Drw.GetFirstToolPath
    For tpIdx = 1 To Drw.GetToolPathCount
        If Not (tp Is Nothing) Then
            If mirrorID <> "" Then
                Set mTool = tp.GetTool
                If Not (mTool Is Nothing) Then
                    If InStr(1, mTool.Name, mirrorID, vbTextCompare) > 0 Then
                        tpCount = tpCount + 1
                        ReDim Preserve collectTP(1 To tpCount)
                        Set collectTP(tpCount) = tp
                    End If
                End If
            Else
                tpCount = tpCount + 1
                ReDim Preserve collectTP(1 To tpCount)
                Set collectTP(tpCount) = tp
            End If
        End If
        Set tp = tp.GetNext
    Next tpIdx
    
    If tpCount = 0 Then
        If mirrorID <> "" Then
            MsgBox "未找到包含 """ & mirrorID & """ 的刀具路径。", vbInformation, "反面镜像"
        End If
        GoTo afterPhase2
    End If
    
    lastop = Drw.Operations.count + 1
    
    ' 按 Sheet 排序
    For Each sh In ni.Sheets
        maxop = 0: minop = lastop
        For Each tp In sh.Paths
            If maxop < tp.OpNo Then maxop = tp.OpNo
            If minop > tp.OpNo Then minop = tp.OpNo
        Next tp
        For I = minop To maxop
            Drw.Operations.Renumber minop, lastop, acamOpADD_TO_OPERATION
        Next I
    Next sh
    
    ' 镜像匹配的刀具路径（保留原路径，后面再删除）
    mirroredCount = 0
    For tpIdx = 1 To tpCount
        Set tp = collectTP(tpIdx)
        If Not (tp Is Nothing) Then
            Set pcopy = tp.CopyTemporary
            If mirrorX Then
                pcopy.MirrorL mirrorVal, miny, mirrorVal, maxy
            Else
                pcopy.MirrorL minx, mirrorVal, maxx, mirrorVal
            End If
            pcopy.Attribute(ATT_IS_REV_SIDE) = 1
            pcopy.Attribute(ATT_PART_NOTE) = "rev"
            pcopy.OpNo = lastop
            lastop = lastop + 1
            pcopy.StoreTemporary
            mirroredCount = mirroredCount + 1
        End If
    Next tpIdx
    
    Drw.Operations.OrderAll
    
    ' 删除正面版件原路径
    For tpIdx = 1 To tpCount
        Set tp = collectTP(tpIdx)
        If Not (tp Is Nothing) Then tp.Delete
    Next tpIdx
    Erase collectTP
    Drw.Operations.OrderAll
    
    If mirroredCount > 0 Then
        Drw.ScreenUpdating = True
        Drw.Redraw
        Drw.ZoomAll
        Drw.Refresh
        DoEvents
    End If
    
afterPhase2:
    
byebye:
    Drw.ScreenUpdating = True
    Drw.Redraw
    Drw.ZoomAll
    
    Set Drw = Nothing
    Set P = Nothing: Set pcopy = Nothing
    Set tp = Nothing
    Set ni = Nothing: Set sh = Nothing
    
    If mirroredCount > 0 Then
        MsgBox "反面镜像完成！" & vbCrLf & vbCrLf & _
               "请按 NC 输出快捷键手动输出代码。", vbInformation, "反面镜像"
    Else
        MsgBox "反面镜像完成！", vbInformation, "反面镜像"
    End If
End Sub
