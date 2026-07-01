' ==============================================================================
' CCC功能 — modMirror 反面镜像
' ==============================================================================
' 算法：
'   1. 镜像 Sheet 几何（复制 + MirrorL + Name/属性标记）
'   2. 镜像 BM 刀具路径到反面版件，删除正面版件原路径
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


' ==============================================================================


' ==============================================================================


' ==============================================================================
Sub 反面镜像()
    fMirror.Show vbModeless
End Sub


' ==============================================================================
' DoMirror — 执行反面镜像（由 fMirror.cmdOK_Click 调用）
' 完全参照 RevNest v1.2 g_AroundX / g_AroundY 算法。
' mirrorX:       True=绕X轴（水平线，垂直翻转），False=绕Y轴（垂直线，水平翻转）
' mirrorID:      刀具名过滤标识（空=不过滤全部镜像，非空=仅匹配刀具名的路径）
' 内部硬编码：bSheetOrder=True（按 Sheet 排序），相同刀具+相同加工方式合并到同一 OP
' ==============================================================================
Public Sub DoMirror(ByVal mirrorX As Boolean, _
                    Optional ByVal mirrorID As String = "")
    
    On Error Resume Next
    
    Dim Drw As Drawing
    Dim P As Path
    Dim pcopy As Path
    Dim lastsheet As Path
    Dim strName As String
    Dim minx As Double, maxx As Double
    Dim miny As Double, maxy As Double
    Dim mirrorVal As Double
    Dim count As Long
    Dim grp As Integer
    Dim elem As Element
    Dim high As Double, big As Double
    Dim exmin As Double, exmax As Double
    Dim eymin As Double, eymax As Double
    Dim textx As Double, texty As Double
    Dim tp As Path
    Dim coll As Paths
    Dim wrd As String
    Dim ni As NestInformation
    Dim sh As NestSheet
    Dim I As Long
    Dim lastop As Long, sheetop As Long
    Dim maxop As Long, minop As Long
    Dim T As Text, T2 As Text
    Dim pT As Path
    Dim psT As Paths
    Dim dblDim As Double
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
        ' 绕垂直轴镜像（X轴方向）：镜像线在 Sheet 左侧 5% 处
        mirrorVal = minx - ((maxx - minx) * 0.05)
    Else
        ' 绕水平轴镜像（Y轴方向）：镜像线在 Sheet 下方 5% 处
        mirrorVal = miny - ((maxy - miny) * 0.05)
    End If
    
    ' ======================================================================
    ' Phase 1 — 镜像 Sheet 几何
    ' ======================================================================
    Set lastsheet = Nothing
    Set P = Drw.GetFirstGeo
    For count = Drw.GetGeoCount To 1 Step -1
        If P.Sheet Or P.Dimension Then
            ' 如果是标注，只处理气泡（bobble），跳过文字标注
            If P.Dimension Then
                If P.Attribute(ATT_IS_BOBBLE) = 0 Then GoTo loopnext
            End If
            
            ' 复制并镜像
            Set pcopy = P.CopyTemporary
            If mirrorX Then
                pcopy.MirrorL mirrorVal, miny, mirrorVal, maxy
            Else
                pcopy.MirrorL minx, mirrorVal, maxx, mirrorVal
            End If
            pcopy.StoreTemporary
            
            ' --- 处理 Sheet ---
            If P.Sheet Then
                Set lastsheet = P
                grp = Drw.GetNextGroupNumberForGeometries
                
                strName = pcopy.Attribute(ATT_SHEET_IDENT)
                strName = strName & " rev"
                
                pcopy.Attribute(ATT_SHEET_IDENT) = strName
                pcopy.Attribute(ATT_SHEET_MATERIAL) = P.Attribute(ATT_SHEET_MATERIAL)
                pcopy.Attribute(ATT_SHEET_THICKNESS) = P.Attribute(ATT_SHEET_THICKNESS)
                pcopy.Attribute(ATT_IS_REV_SIDE) = 1
                pcopy.Name = strName
                
                ' --- 镜像 Sheet 内的文字 ---
                For Each T In Drw.Text
                    If T.Attribute(ATT_REV_TEXT) = 0 Then
                        Set psT = T.ConvertToTemporaryGeometry
                        If psT(1).TestInsidePath(P) = acamResultTRUE Then
                            For Each pT In psT
                                If mirrorX Then
                                    pT.MirrorL mirrorVal, miny, mirrorVal, maxy
                                Else
                                    pT.MirrorL minx, mirrorVal, maxx, mirrorVal
                                End If
                            Next pT
                            
                            If mirrorX Then
                                psT.GetExtentL dblDim, 0, 0, 0
                            Else
                                psT.GetExtentL 0, dblDim, 0, 0
                            End If
                            
                            Set T2 = T.Copy
                            If mirrorX Then
                                T2.MoveL (dblDim - T.MinXL), 0
                            Else
                                T2.MoveL 0, (dblDim - T.MinYL)
                            End If
                            
                            T.Attribute(ATT_REV_TEXT) = 1
                            T2.Attribute(ATT_REV_TEXT) = 1
                            T2.Attribute(ATT_IS_REV_SIDE) = 1
                            T2.Attribute(ATT_SHEET_IDENT) = strName
                        End If
                        psT.Delete
                    End If
                Next T
                
            Else
                ' --- 处理气泡（标注圆）— 写入 Sheet 名称文字 ---
                Set elem = P.GetFirstElem
                If elem.IsArc Then
                    wrd = lastsheet.Attribute(ATT_SHEET_IDENT)
                    wrd = LCase(wrd)
                    high = pcopy.MaxYL - pcopy.MinYL
                    
                    ' 创建临时文字测算尺寸
                    Set coll = Drw.CreateText(wrd, 0, 0, high)
                    exmin = 1E+20: eymin = 1E+20
                    exmax = -1E+20: eymax = -1E+20
                    For Each tp In coll
                        If tp.MinXL < exmin Then exmin = tp.MinXL
                        If tp.MinYL < eymin Then eymin = tp.MinYL
                        If tp.MaxXL > exmax Then exmax = tp.MaxXL
                        If tp.MaxYL > eymax Then eymax = tp.MaxYL
                    Next
                    If (exmax - exmin) > (eymax - eymin) Then
                        big = exmax - exmin
                    Else
                        big = eymax - eymin
                    End If
                    coll.Selected = True
                    Drw.DeleteSelected
                    
                    Dim SF As Double
                    SF = (0.707 * high / big)
                    textx = pcopy.MinXL + ((pcopy.MaxXL - pcopy.MinXL) / 2) - (((exmax - exmin) / 2) * SF)
                    texty = pcopy.MinYL + ((pcopy.MaxYL - pcopy.MinYL) / 2) - (((eymax - eymin) / 2) * SF)
                    high = high * SF
                    
                    Set coll = Drw.CreateText(wrd, textx, texty, high)
                    For Each tp In coll
                        tp.Group = grp
                        tp.Dimension = True
                    Next
                End If
            End If
            
            ' 设置组编号
            pcopy.Group = grp
        End If
        
loopnext:
        Set P = P.GetNext
    Next
    
    ' ======================================================================
    ' Phase 2 — 镜像 BM 刀具路径到反面版件，删除正面版件
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
                ' 没有过滤标识则处理所有刀具路径
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
    
    ' ======================================================================
    ' 缓存每条路径的原始 OpNo 和所属版件名称（在 Renumber 之前执行，
    ' 因为 Renumber 会改变所有 OpNo，导致后续无法区分路径属于哪张板）
    ' ======================================================================
    Dim origOpNos() As Long
    Dim sheetNames() As String
    ReDim origOpNos(1 To tpCount)
    ReDim sheetNames(1 To tpCount)
    
    For tpIdx = 1 To tpCount
        Set tp = collectTP(tpIdx)
        If Not (tp Is Nothing) Then
            origOpNos(tpIdx) = tp.OpNo
            
            ' 在当前 OpNo（原始值）下查找所属版件
            Dim shtName As String: shtName = ""
            Dim s3 As NestSheet
            For Each s3 In ni.Sheets
                Dim tpInSht As Path
                For Each tpInSht In s3.Paths
                    If tpInSht.OpNo = origOpNos(tpIdx) Then
                        shtName = s3.Geometry.Attribute(ATT_SHEET_IDENT)
                        Exit For
                    End If
                Next tpInSht
                If shtName <> "" Then Exit For
            Next s3
            If shtName = "" Then shtName = "未知"
            sheetNames(tpIdx) = shtName
        End If
    Next tpIdx
    
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
    
    ' 修正：Renumber 已将所有路径合并到 lastop，lastop 已成为"垃圾收集操作"。
    ' 在此递增 lastop，使其指向一个干净的空操作号给第一组镜像路径使用。
    lastop = lastop + 1
    
    ' 镜像匹配的刀具路径（保留原路径，后面再删除）
    ' 分组策略：解散原 OP，按加工方式+刀具重新分组
    ' 加工方式从原始 SubOperation 名称提取（"刀具" 前部分）
    mirroredCount = 0
    Dim tgtOp As Long
    Dim keyIdx As Long, keyCount As Long
    Dim keys() As String, opNos() As Long
    keyCount = 0
    
    For tpIdx = 1 To tpCount
        Set tp = collectTP(tpIdx)
        If Not (tp Is Nothing) Then
            ' 获取刀具号
            Set mTool = tp.GetTool
            Dim tNum As Long: tNum = 0
            If Not (mTool Is Nothing) Then tNum = mTool.Number
            
            ' 获取该路径所在的版件名称（从 Renumber 前的预缓存中读取）
            Dim sheetName As String: sheetName = ""
            If tpIdx <= UBound(sheetNames) Then sheetName = sheetNames(tpIdx)
            If sheetName = "" Then sheetName = "未知"
            
            ' 分组键：版件 + 刀具号（加工方式由刀具决定）
            Dim grpKey As String: grpKey = sheetName & "|" & CStr(tNum)
            
            ' 查找分组键
            tgtOp = 0
            For keyIdx = 1 To keyCount
                If keys(keyIdx) = grpKey Then
                    tgtOp = opNos(keyIdx)
                    Exit For
                End If
            Next keyIdx
            ' 没找到则新建
            If tgtOp = 0 Then
                keyCount = keyCount + 1
                ReDim Preserve keys(1 To keyCount)
                ReDim Preserve opNos(1 To keyCount)
                keys(keyCount) = grpKey
                tgtOp = lastop
                opNos(keyCount) = tgtOp
                lastop = lastop + 1
            End If
            
            Set pcopy = tp.CopyTemporary
            If mirrorX Then
                pcopy.MirrorL mirrorVal, miny, mirrorVal, maxy
            Else
                pcopy.MirrorL minx, mirrorVal, maxx, mirrorVal
            End If
            pcopy.Attribute(ATT_IS_REV_SIDE) = 1
            pcopy.OpNo = tgtOp
            pcopy.StoreTemporary
            mirroredCount = mirroredCount + 1
        End If
    Next tpIdx
    
    ' 删除正面版件原路径
    For tpIdx = 1 To tpCount
        Set tp = collectTP(tpIdx)
        If Not (tp Is Nothing) Then tp.Delete
    Next tpIdx
    Erase collectTP
    
    ' 恢复屏幕刷新后再 OrderAll（ScreenUpdating=False 时 OrderAll 不生效）
    Drw.ScreenUpdating = True
    Drw.Redraw
    On Error GoTo 0
    Drw.Operations.OrderAll
    On Error Resume Next
    
    ' 强制刷新 Project Bar 以更新操作名称和顺序显示
    App.Frame.ProjectBarUpdating = False
    App.Frame.ProjectBarUpdating = True
    
    If mirroredCount > 0 Then
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
    Set lastsheet = Nothing: Set elem = Nothing
    Set tp = Nothing: Set coll = Nothing
    Set ni = Nothing: Set sh = Nothing
    Set T = Nothing: Set T2 = Nothing
    Set pT = Nothing: Set psT = Nothing
    
    If mirroredCount > 0 Then
        MsgBox "反面镜像完成！" & vbCrLf & vbCrLf & _
               "请按 NC 输出快捷键手动输出代码。", vbInformation, "反面镜像"
    Else
        MsgBox "反面镜像完成！", vbInformation, "反面镜像"
    End If
End Sub
