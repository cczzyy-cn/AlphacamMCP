' ==============================================================================
' CCC功能 — modMirror 反面镜像
' ==============================================================================
Option Explicit
Option Private Module

Sub 反面镜像()
    Dim drw As Drawing: Set drw = App.ActiveDrawing
    If drw Is Nothing Then MsgBox "没有活动图纸！", vbExclamation, "反面镜像": Exit Sub
    On Error Resume Next
    Dim ni As Object: Set ni = drw.GetNestInformation
    If ni Is Nothing Or ni.Sheets.Count = 0 Then
        MsgBox "当前图纸没有排版信息！" & vbCrLf & "请先进行排版操作。", vbExclamation, "反面镜像"
        Exit Sub
    End If
    If drw.GetToolPathCount = 0 Then
        MsgBox "没有找到刀具路径！", vbExclamation, "反面镜像"
        Exit Sub
    End If
    Dim iChoice As Integer
    iChoice = MsgBox("绕哪个轴镜像？" & vbCrLf & "是(Y) = 绕X轴（垂直镜像）" & vbCrLf & "否(N) = 绕Y轴（水平镜像）", vbYesNo + vbQuestion, "反面镜像")
    App.SetUndoCommandName "反面镜像"
    App.SetUndoPoint
    drw.ScreenUpdating = False
    If iChoice = vbYes Then
        Call MirrorAroundX(drw, ni)
    Else
        Call MirrorAroundY(drw, ni)
    End If
    drw.Operations.OrderAll
    drw.ScreenUpdating = True
    drw.Redraw
    drw.ZoomAll
    MsgBox "反面镜像完成！", vbInformation, "反面镜像"
End Sub

Private Sub MirrorAroundX(drw As Drawing, ni As Object)
    On Error Resume Next
    Dim minx As Double, maxx As Double, miny As Double, maxy As Double
    Dim sh As Object, P As Path, pcopy As Path
    minx = 1E+20: miny = 1E+20
    maxx = -1E+20: maxy = -1E+20
    For Each sh In ni.Sheets
        Set P = sh.Geometry
        If P.MinXL < minx Then minx = P.MinXL
        If P.MinYL < miny Then miny = P.MinYL
        If P.MaxXL > maxx Then maxx = P.MaxXL
        If P.MaxYL > maxy Then maxy = P.MaxYL
    Next sh
    Dim mirrorx As Double: mirrorx = minx - ((maxx - minx) * 0.05)
    Dim grp As Long: grp = drw.GetNextGroupNumberForGeometries
    Dim count As Integer: Set P = drw.GetFirstGeo
    For count = drw.GetGeoCount To 1 Step -1
        If P.Sheet Or P.Dimension Then
            Set pcopy = P.CopyTemporary
            pcopy.MirrorL mirrorx, miny, mirrorx, maxy
            pcopy.StoreTemporary
            pcopy.Group = grp
        End If
        Set P = P.GetNext
    Next
End Sub

Private Sub MirrorAroundY(drw As Drawing, ni As Object)
    On Error Resume Next
    Dim minx As Double, maxx As Double, miny As Double, maxy As Double
    Dim sh As Object, P As Path, pcopy As Path
    minx = 1E+20: miny = 1E+20
    maxx = -1E+20: maxy = -1E+20
    For Each sh In ni.Sheets
        Set P = sh.Geometry
        If P.MinXL < minx Then minx = P.MinXL
        If P.MinYL < miny Then miny = P.MinYL
        If P.MaxXL > maxx Then maxx = P.MaxXL
        If P.MaxYL > maxy Then maxy = P.MaxYL
    Next sh
    Dim mirrory As Double: mirrory = miny - ((maxy - miny) * 0.05)
    Dim grp As Long: grp = drw.GetNextGroupNumberForGeometries
    Dim count As Integer: Set P = drw.GetFirstGeo
    For count = drw.GetGeoCount To 1 Step -1
        If P.Sheet Or P.Dimension Then
            Set pcopy = P.CopyTemporary
            pcopy.MirrorL minx, mirrory, maxx, mirrory
            pcopy.StoreTemporary
            pcopy.Group = grp
        End If
        Set P = P.GetNext
    Next
End Sub
