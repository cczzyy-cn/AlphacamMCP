import os

src = r'C:\Users\C\Desktop\AlphaCAM-MCP\modRamp_old.bas'
dst = r'C:\Users\C\Desktop\AlphaCAM-MCP\CCC功能\modRamp.bas'

content = open(src, 'r', encoding='gbk').read()

# Fix 1: Replace SetGeoStartToSheetSide with unique-offset version
old_sub = '''    Dim mx As Double: mx = (toolGeo.MinXL + toolGeo.MaxXL) / 2
    Dim my As Double: my = (toolGeo.MinYL + toolGeo.MaxYL) / 2
    Dim startX As Double, startY As Double
    Dim bestDist As Double: bestDist = 1E+30
    Dim d As Double

    ' 左边中点
    d = Abs(scx - toolGeo.MinXL) + Abs(scy - my)
    If d < bestDist Then bestDist = d: startX = toolGeo.MinXL: startY = my

    ' 右边中点
    d = Abs(scx - toolGeo.MaxXL) + Abs(scy - my)
    If d < bestDist Then bestDist = d: startX = toolGeo.MaxXL: startY = my

    ' 下边中点
    d = Abs(scx - mx) + Abs(scy - toolGeo.MinYL)
    If d < bestDist Then bestDist = d: startX = mx: startY = toolGeo.MinYL

    ' 上边中点
    d = Abs(scx - mx) + Abs(scy - toolGeo.MaxYL)
    If d < bestDist Then bestDist = d: startX = mx: startY = toolGeo.MaxYL

    toolGeo.SetStartPoint startX, startY
End Sub'''

new_sub = '''    Dim mx As Double: mx = (toolGeo.MinXL + toolGeo.MaxXL) / 2
    Dim my As Double: my = (toolGeo.MinYL + toolGeo.MaxYL) / 2
    Dim startX As Double, startY As Double
    Dim bestDist As Double: bestDist = 1E+30
    Dim d As Double

    d = Abs(scx - toolGeo.MinXL) + Abs(scy - my)
    If d < bestDist Then bestDist = d: startX = toolGeo.MinXL: startY = my
    d = Abs(scx - toolGeo.MaxXL) + Abs(scy - my)
    If d < bestDist Then bestDist = d: startX = toolGeo.MaxXL: startY = my
    d = Abs(scx - mx) + Abs(scy - toolGeo.MinYL)
    If d < bestDist Then bestDist = d: startX = mx: startY = toolGeo.MinYL
    d = Abs(scx - mx) + Abs(scy - toolGeo.MaxYL)
    If d < bestDist Then bestDist = d: startX = mx: startY = toolGeo.MaxYL

    ' 基于几何位置哈希偏移，每块板的起点不同（防重叠）
    Dim uniOff As Double: uniOff = (Abs(toolGeo.MinXL * 10 + toolGeo.MaxYL * 7)) Mod 50
    If startX = toolGeo.MinXL Or startX = toolGeo.MaxXL Then
        startY = startY + uniOff
        If startY > toolGeo.MaxYL - 5 Then startY = toolGeo.MinYL + 5
        If startY < toolGeo.MinYL + 5 Then startY = toolGeo.MinYL + 5
    Else
        startX = startX + uniOff
        If startX > toolGeo.MaxXL - 5 Then startX = toolGeo.MinXL + 5
        If startX < toolGeo.MinXL + 5 Then startX = toolGeo.MinXL + 5
    End If

    toolGeo.SetStartPoint startX, startY
End Sub'''

content = content.replace(old_sub, new_sub)

# Fix 2: Add tag to ramp end - last 10% of ramp rises 0.5mm
old_ramp = '''        ' 斜坡段：沿路径逐步下刀（Z 按实际水平距离比例，保证角度精确）
        Dim s As Long, px As Double, py As Double, pelem As Element
        For s = 1 To steps
            Dim d As Double: d = rampStartDist + POINT_STEP * s
            If d > geoLen Then d = geoLen
            Dim actDist As Double: actDist = d - rampStartDist
            If toolGeo.PointAtDistanceAlongPathL(d, px, py, pelem) Then
                mtp.Add3DLine px, py, -actualDepthAbs * (actDist / sloopDist)
            End If
        Next'''

new_ramp = '''        ' 斜坡段：最后 1mm 上抬 0.5mm 留连接点（tag）
        Dim s As Long, px As Double, py As Double, pelem As Element
        For s = 1 To steps
            Dim d As Double: d = rampStartDist + POINT_STEP * s
            If d > geoLen Then d = geoLen
            Dim actDist As Double: actDist = d - rampStartDist
            Dim ratio As Double: ratio = actDist / sloopDist
            Dim z As Double: z = -actualDepthAbs * ratio
            If ratio > 0.9 Then
                z = z + 0.5 * ((ratio - 0.9) / 0.1)
            End If
            If toolGeo.PointAtDistanceAlongPathL(d, px, py, pelem) Then
                mtp.Add3DLine px, py, z
            End If
        Next'''

content = content.replace(old_ramp, new_ramp)

# Fix 3: Last element at full depth + 0.5mm (tag)
old_full = '''        Dim elems2 As Elements: Set elems2 = toolGeo.Elements
        If Not (elems2 Is Nothing) Then
            For ei = 1 To elems2.Count
                Set elem = elems2(ei)
                If Not (elem Is Nothing) Then
                    If elem.IsLine Then
                        mtp.Add3DLine elem.EndXL, elem.EndYL, finalDepth
                    ElseIf elem.IsArc Then
                        mtp.Add3DArcPointCenter elem.EndXL, elem.EndYL, finalDepth, _
                                                 elem.CenterXL, elem.CenterYL, elem.CW
                    End If
                End If
            Next ei
        End If'''

new_full = '''        Dim elems2 As Elements: Set elems2 = toolGeo.Elements
        If Not (elems2 Is Nothing) Then
            For ei = 1 To elems2.Count
                Set elem = elems2(ei)
                If Not (elem Is Nothing) Then
                    Dim z As Double: z = finalDepth
                    If ei = elems2.Count Then z = finalDepth + 0.5
                    If elem.IsLine Then
                        mtp.Add3DLine elem.EndXL, elem.EndYL, z
                    ElseIf elem.IsArc Then
                        mtp.Add3DArcPointCenter elem.EndXL, elem.EndYL, z, _
                                                 elem.CenterXL, elem.CenterYL, elem.CW
                    End If
                End If
            Next ei
        End If'''

content = content.replace(old_full, new_full)

open(dst, 'w', encoding='gbk').write(content)
print('Fixed version written')
