Sub TestDoor()
    Dim drw As Drawing
    Set drw = App.ActiveDrawing
    If drw Is Nothing Then Exit Sub
    
    Dim cp As Path
    Set cp = drw.GetFirstGeo
    Do While Not (cp Is Nothing)
        cp.Erase
        Set cp = drw.GetFirstGeo
    Loop
    
    ' 参数
    Dim dblWidth As Double: dblWidth = 600
    Dim dblLength As Double: dblLength = 1200
    Dim dblStile As Double: dblStile = 60
    Dim dblRail As Double: dblRail = 60
    Dim dblStileLeft As Double: dblStileLeft = 60
    Dim dblRailBottom As Double: dblRailBottom = 60
    Dim dblArchHeight As Double: dblArchHeight = 50
    Dim dblLeg As Double: dblLeg = 20
    Dim dblBlend As Double: dblBlend = 25
    Dim dblInnerOffset As Double: dblInnerOffset = 20
    Dim dblGridSpacing As Double: dblGridSpacing = 50
    Dim dblLowerRectHeight As Double: dblLowerRectHeight = 800
    Dim dblMiddleSpacing As Double: dblMiddleSpacing = 55
    
    Dim gb As Double
    gb = dblRailBottom
    If dblLowerRectHeight > 0 Then
        gb = dblRailBottom + dblLowerRectHeight + dblMiddleSpacing
    End If
    
    Dim le As Double: le = dblStileLeft
    Dim re As Double: re = dblWidth - dblStile
    Dim cx As Double: cx = (le + re) / 2
    
    Set cp = drw.CreateRectangle(0, 0, dblWidth, dblLength)
    
    ' 圆弧
    Dim geo As Geo2D, p1 As Path
    Set geo = drw.Create2DGeometry(le, gb)
    With geo
        .AddLine le, (dblLength - (dblRail + dblArchHeight))
        .AddLine (le + dblLeg + (dblBlend / 2)), (dblLength - (dblRail + dblArchHeight))
        .AddArc2Point (le + ((dblWidth - (dblStile + le)) / 2)), (dblLength - dblRail), (dblWidth - (dblStile + dblLeg + (dblBlend / 2))), (dblLength - (dblRail + dblArchHeight))
        .AddLine re, (dblLength - (dblRail + dblArchHeight))
        .AddLine re, gb
        Set p1 = .CloseAndFinishLine
    End With
    p1.Fillet dblBlend
    
    ' 下矩形
    Dim p2 As Path
    If dblLowerRectHeight > 0 Then
        Dim geo2 As Geo2D
        Set geo2 = drw.Create2DGeometry(le, dblRailBottom)
        With geo2
            .AddLine le, (dblRailBottom + dblLowerRectHeight)
            .AddLine re, (dblRailBottom + dblLowerRectHeight)
            .AddLine re, dblRailBottom
            Set p2 = .CloseAndFinishLine
        End With
    End If
    
    ' 内偏移
    Dim po As Path
    If dblInnerOffset > 0 Then
        Dim ops As Paths
        Set ops = p1.Offset(dblInnerOffset, 1)
        If Not (ops Is Nothing) Then
            Dim oi As Long
            For oi = 1 To ops.Count
                Set po = ops(oi)
            Next oi
        End If
    End If
    
    ' 网格线
    If dblGridSpacing > 0 Then
        Dim b(1 To 4) As Path
        Set b(1) = drw.Create2DLine(cx, gb, re, gb + (re - cx))
        Set b(2) = drw.Create2DLine(cx, gb, le, gb + (cx - le))
        Set b(3) = drw.Create2DLine(cx, gb, re, gb - (re - cx))
        Set b(4) = drw.Create2DLine(cx, gb, le, gb - (cx - le))
        Dim ni As Long: ni = 1
        Do While True
            Dim dy As Double: dy = ni * dblGridSpacing
            If gb + dy > dblLength Then Exit Do
            Dim ii As Long
            For ii = 1 To 4
                Dim cc As Path
                Set cc = b(ii).Copy
                cc.MoveG 0, dy, 0
            Next ii
            ni = ni + 1
        Loop
    End If
    
    ' 裁剪
    If dblInnerOffset > 0 And dblGridSpacing > 0 Then
        If Not (po Is Nothing) Then
            Dim gx As Path
            Set gx = drw.GetFirstGeo
            Dim guard As Long: guard = 0
            Do While Not (gx Is Nothing) And guard < 5000
                guard = guard + 1
                If Not gx.Closed Then
                    If gx.TestIntersectPath(po, 0, 0) Then
                        drw.SetGeosSelected False
                        po.Selected = True
                        Dim segs As Paths
                        Set segs = gx.BreakWithCuttingGeos
                        drw.SetGeosSelected False
                        If Not (segs Is Nothing) Then
                            Dim si As Long
                            For si = 1 To segs.Count
                                Dim seg As Path
                                Set seg = segs(si)
                                Dim hl As Double: hl = seg.Length / 2
                                Dim hx As Double, hy As Double, hz As Double
                                Dim he As Element
                                If seg.PointAtDistanceAlongPathG(hl, hx, hy, hz, he) Then
                                    If Not po.IsPointInside(hx, hy) Then
                                        seg.Erase
                                    End If
                                End If
                            Next si
                        End If
                    Else
                        gx.Erase
                    End If
                End If
                Set gx = drw.GetFirstGeo
            Loop
            drw.SetGeosSelected False
        End If
    End If
    
    drw.ZoomAll
    MsgBox "测试完成！", vbInformation
End Sub
