---
name: alphacam-coding
description: Generate correct AlphaCAM 2016 R1 VBA code — geometry, toolpaths, machining, events, menus, and CDM/ReverseNest patterns.
---

# Skill: alphacam-coding
> Generate correct VBA code for AlphaCAM 2016 R1 automation — geometry, toolpaths, machining, menus, events, and add-ins.

## Core Rules

1. **Entry point**: Inside AlphaCAM, `App` is the global Application object.
2. **Drawing**: `App.ActiveDrawing` gets the current drawing.
3. **Error handling**: Use `On Error Resume Next` / `On Error GoTo 0`.
4. **All code goes in a Module called `Events`** for add-in event handlers.

## Object Hierarchy
Application (App) -> ActiveDrawing (Drawing) -> Geometries/ToolPaths/Operations/Layers

## Key Drawing Methods
- `Create2DLine(X1,Y1,X2,Y2) As Path`
- `CreateCircle(Diameter, XCen, YCen) As Path`
- `CreateRectangle(X1,Y1,X2,Y2) As Path`
- `Create2DGeometry(X,Y) As Geo2D` (then .AddLine/.AddArc/.CloseAndFinishLine)
- `CreateFastGeometry As FastGeometry` (then .Point/.KnownArc/.Finish)
- `GetFirstGeo` / `Geometries` (collection iteration)
- `UserSelectOneGeo(Prompt)`, `UserSelectMultiGeos(Prompt, Mode)`
- `OutputNC(FileName, OutputTo, VisibleOnly)`
- `ZoomAll`, `Clear(...)`, `Save`, `SaveAs(FileName)`

## Path Properties & Methods
`Selected` (Boolean), `ToolInOut` (acamINSIDE/OUTSIDE/ON_CENTER), `Group` (Long), `Erase`, `Copy`, `MoveG(Dx,Dy,Dz)`, `MirrorG`, `RotateG`, `ScaleG`, `Fillet(Radius)`, `Length`, `Closed`, `Offset(Distance, Direction) As Paths`, `Color`, `GetFirstElem`/`GetLastElem` -> Element with `StartXG`/`StartYG`/`EndXG`/`EndYG`, `TestIntersectPath(Path)`, `IsPointInside(X,Y)`, `TrimWithCuttingGeos(X,Y)`, `BreakWithCuttingGeos() As Paths`, `PointAtDistanceAlongPathG(Dist, X, Y, Z, Element)`, `IntersectWithLine(X1,Y1,X2,Y2,Extend,XInt,YInt) As Long`

## Element Types
- `ElemType` property: `acamLINE`, `acamARC`, `acamCURVE`
- For lines: `StartXG`/`StartYG` (start point), `EndXG`/`EndYG` (end point)
- For arcs: also has `ArcCenterXG`/`ArcCenterYG`/`ArcRadius`/`ArcStartAngle`/`ArcSweepAngle`

## Process Type Constants
`acamProcessROUGH_FINISH=1`, `acamProcessCONTOUR_POCKET=2`, `acamProcessMANUAL=3`, `acamProcessENGRAVE=10`, `acamProcessMACHINE_SURFACE=11`, `acamProcessMACHINE_POLYLINE=12`, `acamProcessDRILL=21`, `acamProcessPECK=22`, `acamProcessTAP=23`, `acamProcessBORE=24`

## Tool Compensation
`acamINSIDE`, `acamOUTSIDE`, `acamON_CENTER`

## CDM Pattern (AdoorMain callback)
```vb
Public Sub AdoorMain(RequiredData As Object, User_1..7 As Variant)
    ' RequiredData: .width, .Length, .CornerRadius, .UserVariables (Collection)
    ' Return paths via: RequiredData.PathsToReturn.Add path
    ' User variables indexed by enum starting at 1
End Sub
```

## Fillet + Offset Pattern
```vb
' Create geometry, apply fillet, offset inward
Set geo = App.ActiveDrawing.Create2DGeometry(x, y)
With geo
    .AddLine ...
    .AddArc2Point ...
    Set pth = .CloseAndFinishLine
End With
pth.Fillet Radius
Dim offsetPaths As Paths
Set offsetPaths = pth.Offset(OffsetDist, 1)  ' 1 = inward
```

### IntersectWithLine -- boundary clipping grid lines (recommended)

Recommended grid line clipping method using `IntersectWithLine` to calculate intersections with a boundary path, no `TrimWithCuttingGeos` needed:

```vb
Dim clipPath As Path
Set clipPath = pthOffset  ' boundary path (closed)

' For each 45 deg line, calc intersections
Dim X1, Y1, X2, Y2 As Double
X1 = innerLeft: Y1 = rBottom - w + dy
X2 = innerRight: Y2 = rBottom + dy

Dim XInt As Variant, YInt As Variant
Dim interCnt As Long
interCnt = clipPath.IntersectWithLine(X1, Y1, X2, Y2, True, XInt, YInt)

Dim gotPts As Boolean: gotPts = False
Dim P1x, P1y, P2x, P2y As Double

If interCnt = 2 Then
    If YInt(0) < YInt(1) Then
        P1x = XInt(0): P1y = YInt(0): P2x = XInt(1): P2y = YInt(1)
    Else
        P1x = XInt(1): P1y = YInt(1): P2x = XInt(0): P2y = YInt(0)
    End If
    gotPts = True
ElseIf interCnt = 1 Then
    If clipPath.IsPointInside(X1, Y1) Then
        P1x = X1: P1y = Y1: P2x = XInt(0): P2y = YInt(0)
    ElseIf clipPath.IsPointInside(X2, Y2) Then
        P1x = XInt(0): P1y = YInt(0): P2x = X2: P2y = Y2
    End If
    gotPts = True
Else
    Dim hx, hy As Double
    hx = (X1 + X2) / 2: hy = (Y1 + Y2) / 2
    If clipPath.IsPointInside(hx, hy) Then
        P1x = X1: P1y = Y1: P2x = X2: P2y = Y2: gotPts = True
    End If
End If

If gotPts Then
    Dim lineSeg As Path
    Set lineSeg = App.ActiveDrawing.Create2DLine(P1x, P1y, P2x, P2y)
    RequiredData.PathsToReturn.Add lineSeg
    
    ' Mirror across center
    Dim lineMir As Path
    Set lineMir = App.ActiveDrawing.Create2DLine(dblWidth-P1x, P1y, dblWidth-P2x, P2y)
    RequiredData.PathsToReturn.Add lineMir
End If
```

**Advantages**: Pure math, no drawing pollution, no CDM interference. Handles 2/1/0 intersection cases fully.

## Nesting API (NestInformation)

```vb
Dim ni As NestInformation: Set ni = Drw.GetNestInformation
' ni.Sheets → NestSheet collection
' ni.Sheets(1).Geometry → Sheet boundary Path
' ni.Sheets(1).Paths → Paths on this sheet
' ni.Sheets(1).Parts → NestPartInstance collection
'   inst.Paths(1) → first path of part (carries ATT_PATH_FILE etc.)
```

## Temporary Path Operations (CopyTemporary / StoreTemporary)

For bulk geometry operations without interfering with the drawing database:

```vb
Set pcopy = P.CopyTemporary   ' Copy to temp buffer
pcopy.MirrorL x1, y1, x2, y2  ' Transform in temp
pcopy.StoreTemporary           ' Write back to database
```

Always iterate geometries **in reverse** when inserting new ones:

```vb
For count = Drw.GetGeoCount To 1 Step -1
    ' ... CopyTemporary → Transform → StoreTemporary
    Set P = P.GetNext
Next
```

## Path Attribute System

Attributes store add-in data on Path objects. Naming: `Company_Country_Initials_Project_Identifier`.

```vb
' Write
P.Attribute("LicomUKsab_nest_path_file") = "C:\parts\door.amd"
' Read (returns Variant)
Dim sName As String: sName = P.Attribute("LicomUKsab_nest_path_file")
' Check attribute with numeric comparison
If P.Attribute("LicomUKsab_is_bobble") = 0 Then ...
```

## Mirroring Pattern (Reverse-Side Nesting)

### Mirror axis calculation
```vb
' Find sheet extents
For Each sh In ni.Sheets
    Set P = sh.Geometry
    If P.MinXL < minx Then minx = P.MinXL
    If P.MaxXL > maxx Then maxx = P.MaxXL
    ' ... similarly miny/maxy
Next sh

' Mirror line 5% outside sheets
mirrorX = minx - ((maxx - minx) * 0.05)  ' Vertical axis (g_AroundX)
mirrorY = miny - ((maxy - miny) * 0.05)  ' Horizontal axis (g_AroundY)
```

### Path transformation order (critical!)
```vb
' 1. Shift (local offset) — before reflection/rotation
curPth.MoveL ShiftX, ShiftY
' 2. Reflect (if part is mirrored) — before rotation
If intReflect = 1 Then
    curPth.MirrorL 0, 1, 0, 0
    rotate = -rotate  ' Reverse rotation direction
End If
' 3. Rotate
curPth.RotateL rotate, 0, 0
' 4. MoveBy (global positioning) — after local transforms
curPth.MoveL MoveX, MoveY
' 5. Mirror to reverse side (global axis)
curPth.MirrorL mirrorX, minY, mirrorX, maxY  ' Around X
' OR
curPth.MirrorL minX, mirrorY, maxX, mirrorY  ' Around Y
```

### Loading reverse-side part files
```vb
strName = P.Attribute("LicomUKsab_nest_path_file")
prefix = Left(strName, Len(strName) - 4)
suffix = Right(strName, 4)
strName = prefix & "_rev" & suffix  ' door_left.amd → door_left_rev.amd
Set tmpdrw = App.OpenTempDrawing(strName)

' TFS#80910: Check for workplanes before inserting
If tmpdrw.WorkPlanes.count > 0 Then
    MsgBox strName & " contains workplanes and cannot be used"
    GoTo skipPart
End If
```

### Operation renumbering
```vb
' Move operations to end for sheet ordering
Drw.Operations.Renumber minOp, targetOp, acamOpADD_TO_OPERATION
' Reorder all operations by OpNo
Drw.Operations.OrderAll
```

## Copying attributes between objects
```vb
' Copy all nest-related attributes from original to mirrored path
curPth.Attribute(ATT_PATH_FILE) = sName
curPth.Attribute(ATT_FIRST_PATH) = isFirstTP
curPth.Attribute(ATT_REQUIRED) = orgPath.Attribute(ATT_REQUIRED)
curPth.Attribute(ATT_NEST_ITEM_NUM) = orgPath.Attribute(ATT_NEST_ITEM_NUM)
curPth.Attribute(ATT_IS_REV_SIDE) = 1
```

## Analyzing compiled VBA (.amb/.arb) add-ins
```bash
# Install oletools
pip install oletools
# Decompile VBA from OLE2 compound document
python3 -m oletools.olevba ReverseNest.amb > output.txt
```

The `.amb` / `.arb` / `.asb` files are OLE2 compound documents containing compressed VBA streams. `olevba` extracts and decompiles them to readable source code.
