' ==============================================================================
' CCC功能 — modMirror 反面镜像（RevNest 完整算法）
' ==============================================================================
' 实现 RevNest v1.2 的 g_AroundX / g_AroundY 镜像逻辑：
'   1. 镜像 Sheet 几何（复制 + MirrorL + 属性标记）
'   2. 加载 _rev 反面零件图档，应用 Nesting 变换（Shift→Reflect→Rotate→Move→Mirror）
'   3. 处理操作编号排序
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
Sub 反面镜像()
    fMirror.Show vbModeless
End Sub


' ==============================================================================
' DoMirror — 执行反面镜像（由 fMirror.cmdOK_Click 调用）
' 完全参照 RevNest v1.2 g_AroundX / g_AroundY 算法。
' mirrorX:       True=绕X轴（水平线，垂直翻转），False=绕Y轴（垂直线，水平翻转）
' mirrorID:      刀具名过滤标识（空=不过滤全部镜像，非空=仅匹配刀具名的路径）
' 内部硬编码：bSheetOrder=True（按 Sheet 排序），bMinToolChanges=False，bIncludeGeos=False
' ==============================================================================
Public Sub DoMirror(ByVal mirrorX As Boolean, _
                    Optional ByVal mirrorID As String = "")
    
    On Error Resume Next
    
    Dim Drw As Drawing
    Dim tmpdrw As Drawing
    Dim P As Path
    Dim pcopy As Path
    Dim newp As Path
    Dim lastsheet As Path
    Dim prefix As String
    Dim suffix As String
    Dim strName As String
    Dim minx As Double, maxx As Double
    Dim miny As Double, maxy As Double
    Dim mirrorVal As Double
    Dim xmove As Double, ymove As Double
    Dim rotate As Double
    Dim count As Long, count2 As Long
    Dim reflect As Integer
    Dim grp As Integer
    Dim isfirst As Long
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
    Dim inst As NestPartInstance
    Dim I As Long
    Dim lastop As Long, sheetop As Long
    Dim maxop As Long, minop As Long
    Dim T As Text, T2 As Text
    Dim pT As Path
    Dim psT As Paths
    Dim dblDim As Double
    Dim mTool As MillTool
    Dim filterTP As Boolean
    
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
    ' Phase 2 — 镜像刀具路径（加载 _rev 文件，应用变换）
    ' ======================================================================
    lastop = Drw.Operations.count + 1
    For Each sh In ni.Sheets
        ' 按 Sheet 排序，将当前 Sheet 的操作移至末尾
        maxop = 0: minop = lastop
        For Each tp In sh.Paths
            If maxop < tp.OpNo Then maxop = tp.OpNo
            If minop > tp.OpNo Then minop = tp.OpNo
        Next tp
        For I = minop To maxop
            Drw.Operations.Renumber minop, lastop, acamOpADD_TO_OPERATION
        Next I
        
        sheetop = lastop
        
        For Each inst In sh.Parts
            Set P = inst.Paths(1)
            If P.Attribute(ATT_FIRST_PATH) = 0 Then GoTo loopagain
            
            ' 获取零件文件路径，添加 _rev 后缀
            strName = P.Attribute(ATT_PATH_FILE)
            prefix = Left(strName, Len(strName) - 4)
            suffix = Right(strName, 4)
            strName = prefix + "_rev" + suffix
            
            ' 加载反面零件图档
            Set tmpdrw = Nothing
            Set tmpdrw = App.OpenTempDrawing(strName)
            
            If Not tmpdrw Is Nothing Then
                ' 检查是否有工作平面
                If tmpdrw.WorkPlanes.count > 0 Then
                    MsgBox strName & " 包含工作平面，无法使用。", vbExclamation, "反面镜像"
                    GoTo loopagain
                End If
                
                ' --- 遍历刀具路径，应用 Nesting 变换 ---
                Set newp = tmpdrw.GetFirstToolPath
                isfirst = 1
                reflect = P.Attribute(ATT_PART_MIRRORED)
                
                For count2 = 1 To tmpdrw.GetToolPathCount
                    ' --- 刀具名过滤（RevNest 功能扩展） ---
                    filterTP = True
                    If mirrorID <> "" Then
                        Set mTool = newp.GetTool
                        If Not (mTool Is Nothing) Then
                            If InStr(1, mTool.Name, mirrorID, vbTextCompare) = 0 Then
                                filterTP = False
                            End If
                        Else
                            filterTP = False
                        End If
                    End If
                    
                    If filterTP Then
                    ' 应用 Shift → Reflect → Rotate → Move → Mirror 变换
                    xmove = P.Attribute(ATT_PART_MOVEX)
                    ymove = P.Attribute(ATT_PART_MOVEY)
                    rotate = P.Attribute(ATT_PART_ROTANGLE)
                    
                    Dim MoveX As Double, MoveY As Double
                    Dim ShiftX As Double, ShiftY As Double
                    MoveX = P.Attribute(ATT_PART_MOVE_BY_X)
                    MoveY = P.Attribute(ATT_PART_MOVE_BY_Y)
                    ShiftX = P.Attribute(ATT_PART_SHIFT_X)
                    ShiftY = P.Attribute(ATT_PART_SHIFT_Y)
                    
                    newp.MoveL ShiftX, ShiftY
                    
                    If reflect = 1 Then
                        newp.MirrorL 0, 1, 0, 0
                        rotate = -rotate
                    End If
                    
                    newp.RotateL rotate, 0, 0
                    newp.MoveL MoveX, MoveY
                    
                    If mirrorX Then
                        newp.MirrorL mirrorVal, miny, mirrorVal, maxy
                    Else
                        newp.MirrorL minx, mirrorVal, maxx, mirrorVal
                    End If
                    
                    ' 设置属性
                    newp.Attribute(ATT_FIRST_PATH) = isfirst
                    newp.Attribute(ATT_PATH_FILE) = strName
                    newp.Attribute(ATT_REQUIRED) = P.Attribute(ATT_REQUIRED)
                    newp.Attribute(ATT_NEST_ITEM_NUM) = P.Attribute(ATT_NEST_ITEM_NUM)
                    newp.Attribute(ATT_IS_REV_SIDE) = 1
                    
                    If isfirst = 1 Then isfirst = 0
                    
                    ' 复制到主图纸，设置操作编号
                    Set pcopy = newp.CopyTemporary
                    
                    pcopy.OpNo = lastop
                    lastop = lastop + 1
                    
                    pcopy.StoreTemporary
                    End If  ' filterTP
                    Set newp = newp.GetNext
                Next count2
                
                ' 不镜像几何路径
            End If
            
loopagain:
        Next inst
    Next sh
    
    Drw.Operations.OrderAll
    
byebye:
    Drw.ScreenUpdating = True
    Drw.Redraw
    
    Set Drw = Nothing: Set tmpdrw = Nothing
    Set P = Nothing: Set pcopy = Nothing: Set newp = Nothing
    Set lastsheet = Nothing: Set elem = Nothing
    Set tp = Nothing: Set coll = Nothing
    Set ni = Nothing: Set sh = Nothing: Set inst = Nothing
    Set T = Nothing: Set T2 = Nothing
    Set pT = Nothing: Set psT = Nothing
    
    MsgBox "反面镜像完成！", vbInformation, "反面镜像"
End Sub
