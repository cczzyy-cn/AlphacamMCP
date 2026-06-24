in file: C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\000\StartUp\Utils\ReverseNest\ReverseNest.amb - OLE stream: 'vao/The VBA Project/_VBA_Project/VBA/modAcam'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

' >< ENUMS ><
'
Public Enum AlphaModuleType
        alphaMod_WIRE = 69          'E
        alphaMod_PROFILING = 76     'L
        alphaMod_MILL = 77          'M
        alphaMod_ROUTER = 82        'R
        alphaMod_STONE = 83         'S
        alphaMod_LATHE = 84         'T
End Enum

Public Enum AlphaVariableType
        alphaVarType_STRING = 0
        alphaVarType_INTEGER
        alphaVarType_LONG
        alphaVarType_SINGLE
        alphaVarType_DOUBLE
        alphaVarType_BOOLEAN
End Enum

Public Enum AlphaSpecialKey
        alphaSpecialKey_NORMAL = 0
        alphaSpecialKey_DEALER = 1
        alphaSpecialKey_EDUC = 2
        alphaSpecialKey_COUNTDOWN = 4
        alphaSpecialKey_DEALER_COUNTDOWN = 5
End Enum

Public Enum AlphaFileType
        alphaFile_DRAWING = 1
        alphaFile_NC = 2
        alphaFile_POST = 3
        alphaFile_TOOL = 4
        alphaFile_FONT = 5
        alphaFile_THREAD = 6
        alphaFile_MACRO = 7
        alphaFile_NESTLIST = 8
        alphaFile_VBAPROJECT = 9
        alphaFile_STYLE = 10
        alphaFile_MACHINECONFIG = 11
        alphaFile_TEMPLATE = 12
End Enum

Public Enum AlphaExtensionType
        alphaExtension_DRAWING = 0
        alphaExtension_TOOL
        alphaExtension_FONT
        alphaExtension_POST
        alphaExtension_VBA
        alphaExtension_MATERIAL
        alphaExtension_NC
        alphaExtension_NESTLIST
        alphaExtension_DXF
        alphaExtension_DWG
        alphaExtension_IGES
        alphaExtension_EMF
        alphaExtension_WMF
        alphaExtension_DAT
        alphaExtension_TEXT
        alphaExtension_STYLE
        alphaExtension_MACHINECONFIG
        alphaExtension_VBAPOSTCONFIG
        alphaExtension_TEMPLATE
End Enum

Public Enum AlphaWorkFace
        alphaWorkFace_UNKNOWN = -1
        alphaWorkFace_TOP = 0
        alphaWorkFace_FRONT = 1
        alphaWorkFace_RIGHT = 2
        alphaWorkFace_BACK = 3
        alphaWorkFace_LEFT = 4
        alphaWorkFace_BOTTOM = 5
End Enum

Public Enum AlphaDragObjects
        alphaDrag_GEOS = 1
        alphaDrag_TOOLPATHS = 2
        alphaDrag_BOTH = 4
End Enum

Public Enum AlphaLineOrientation
        alphaLine_ANGLED
        alphaLine_HORIZONTAL
        alphaLine_VERTICAL
End Enum

Public Enum AlphaEndToExtend
        alphaExtend_START = 1
        alphaExtend_END = 2
End Enum

Public Enum AlphaQuadrant
        alphaQuadrant_I = 0
        alphaQuadrant_II = 1
        alphaQuadrant_III = 2
        alphaQuadrant_IV = 3
End Enum

Public Enum AlphaSetUnsetOpenElementMethod
        alphaOpenE_AUTO = 0
        alphaOpenE_SET = 1
        alphaOpenE_UNSET = 2
End Enum

Public Enum AlphaNestExtension
        alphaNestExt_CUSTOM = -1
        alphaNestExt_CUT_HOLE_PART = 1
        alphaNestExt_SUPPRESS_FINAL_SORT
        alphaNestExt_GROUP_EACH_PART
        alphaNestExt_ROTATED_PARTS_FIRST
        alphaNestExt_SUPPRESS_REDRAW
        alphaNestExt_REMOVE_GROUPS
        alphaNestExt_CUT_SMALL_FIRST
        alphaNestExt_QUANTITY_MERGE
        alphaNestExt_QUANTITY_MULTIPLIER
        alphaNestExt_ONION_SKIN
        alphaNestExt_NEST_SMALL_TO_LARGE
End Enum

Public Enum AlphaProgramLevel
        alphaLevel_CAD = -1
        alphaLevel_BASIC = 1
        alphaLevel_STANDARD = 2
        alphaLevel_ADVANCED = 3
        alphaLevel_ADVANCED3D3AXIS = 4
        alphaLevel_ADVANCED3D5AXIS = 5
        alphaLevel_VIEWPLUS = 6
        alphaLevel_PROUTAPS = 7
        alphaLevel_OEM5AXIS = 8
End Enum

Public Enum AlphaViewAnimationSpeed
        alphaViewAnimationSpeed_SLOW = 0
        alphaViewAnimationSpeed_MEDIUM = 1
        alphaViewAnimationSpeed_FAST = 2
End Enum

' >< UDTs ><
'
Public Type SAW_MILLDATA_PROPERTIES
        SawInternalCorners          As AcamSawCornerType
        SawExternalCorners          As AcamSawCornerType
        SawHeadPosition             As AcamSawHeadPosition
        SawInternalCornerDistance   As Double
        SawExternalCornerDistance   As Double
        SawAngle                    As Double
        SawIncludeArcs              As Boolean
        SawMinimumArcRadius         As Double
        SawResetAngles              As Boolean
End Type

' >< CONSTANTS ><
'
Private Const WM_SETREDRAW              As Long = &HB

Private Const LicomUKDMBMCType          As String = "LicomUKDMBMCType"
Private Const LicomUKSAJFixtureLayer    As String = "LicomUKSAJFixtureLayer"
Private Const LicomUKDMBGeoZLevelTop    As String = "LicomUKDMBGeoZLevelTop"
Private Const LicomUKDMBGeoZLevelBottom As String = "LicomUKDMBGeoZLevelBottom"
Private Const LicomUKDMBBmpName         As String = "LicomUKDMBBmpName"

Private Const LicomUKDMBSawAngle        As String = "LicomUKDMBSawAngle"
Private Const LicomUKDMBResetAngles     As String = "_LicomUKDMBResetAngles"
Private Const LicomUKDMBSaw1            As String = "_LicomUKDMBSaw1"
Private Const LicomUKDMBSaw2            As String = "_LicomUKDMBSaw2"
Private Const LicomUKDMBSaw3            As String = "_LicomUKDMBSaw3"
Private Const LicomUKDMBSaw4            As String = "_LicomUKDMBSaw4"
Private Const LicomUKDMBSaw5            As String = "_LicomUKDMBSaw5"
Private Const LicomUKDMBSawDeepestCut   As String = "_LicomUKDMBSawDeepestCut"
Private Const LicomUKDMBSawIncludeArcs  As String = "_LicomUKDMBSawIncludeArcs"
Private Const LicomUKDMBSawMinRad       As String = "_LicomUKDMBSawMinRad"
'

Public Sub g_GetViewAnimationSettings(Optional bViewPoint As Boolean, Optional bZoom As Boolean, Optional iSpeed As AlphaViewAnimationSpeed)
        
        Dim strMod                  As String
        Dim strKey                  As String
        
On Error Resume Next
        
        strMod = gs_RawModuleName
        
        If (Len(Trim$(strMod)) = 0) Then Exit Sub
        
        strKey = "Software\LicomSystems\" & strMod & "\Settings"
        
        bViewPoint = CBool(gs_ReadRegKey(strKey, "AnimateView", , "0"))
        bZoom = CBool(gs_ReadRegKey(strKey, "AnimateZoom", , "0"))
        iSpeed = CInt(gs_ReadRegKey(strKey, "AnimateSpeed", , "1"))

End Sub

Public Sub g_SetViewAnimationSettings(ByVal bViewPoint As Boolean, ByVal bZoom As Boolean, ByVal iSpeed As AlphaViewAnimationSpeed)
        
        Dim strMod                  As String
        Dim strKey                  As String
        Dim blnRet                  As Boolean
        
On Error Resume Next
        
        strMod = gs_RawModuleName
        
        If (Len(Trim$(strMod)) = 0) Then Exit Sub
        
        strKey = "Software\LicomSystems\" & strMod & "\Settings"
        
        blnRet = gb_WriteRegKey(REG_DWORD, strKey, "AnimateView", Abs(bViewPoint))
        blnRet = gb_WriteRegKey(REG_DWORD, strKey, "AnimateZoom", Abs(bZoom))
        blnRet = gb_WriteRegKey(REG_DWORD, strKey, "AnimateSpeed", iSpeed)
        
End Sub

Public Sub g_GetExtentsG(PS As Paths, uMin As POINT_XYZ, uMax As POINT_XYZ)
        
        Dim P                       As Path
        Dim udtMin                  As POINT_XYZ
        Dim udtMax                  As POINT_XYZ
        Dim udtMin2                 As POINT_XYZ
        Dim udtMax2                 As POINT_XYZ
        
        If Not (PS Is Nothing) Then
                    
                With udtMin2
                        .X = 999999999
                        .Y = 999999999
                End With
                
                With udtMax2
                        .X = -999999999
                        .Y = -999999999
                End With
    
                For Each P In PS
                                                
                        If P.GetFeedExtent(udtMin.X, udtMin.Y, udtMax.X, udtMax.Y) Then
                                
                                Call g_LtoG(P.GetWorkPlane, udtMin.X, udtMin.Y, 0)
                                Call g_LtoG(P.GetWorkPlane, udtMax.X, udtMax.Y, 0)
                                
                                If (udtMin.X <= udtMin2.X) Then udtMin2.X = udtMin.X
                                If (udtMin.Y <= udtMin2.Y) Then udtMin2.Y = udtMin.Y
                                If (udtMax.X >= udtMax2.X) Then udtMax2.X = udtMax.X
                                If (udtMax.Y >= udtMax2.Y) Then udtMax2.Y = udtMax.Y

                        End If
                
                Next P
        
        End If
        
        uMin.X = udtMin2.X
        uMin.Y = udtMin2.Y
        uMax.X = udtMax2.X
        uMax.Y = udtMax2.Y
                
        Set P = Nothing
        
End Sub

Public Sub g_GetWAAandWACandWTC(WP As WorkPlane, WAA As Double, WAC As Double, WTC As Double)

        Dim X                       As Double
        Dim Y                       As Double
        Dim Z                       As Double
        Dim dWAA                    As Double
        Dim dWAC                    As Double
        Dim dWTC                    As Double

        ' check that a workplane has actually been sent to the sub
        If WP Is Nothing Then Exit Sub

        ' get the X, Y & Z figures for the selected workplane
        X = WP.Tmat(8)
        Y = WP.Tmat(9)
        Z = WP.Tmat(10)

        ' calculate waa
        If (X = 0) And (Y = 0) Then
                dWAA = 0
        Else
                dWAA = CDbl(Format$(App.Frame.Evaluate("atan2(" & X & ", " & Y & ")"), "0.####"))
        End If

        ' get the X, Y & Z figures for the selected workplane
        X = WP.Tmat(2)
        Y = WP.Tmat(6)
        Z = WP.Tmat(10)

        ' calculate wac
        If (X = 0) And (Y = 0) Then
                dWAC = 0
        Else
                dWAC = CDbl(Format$(App.Frame.Evaluate("atan2(" & Y & ", " & X & ")"), "0.####"))
        End If

        ' calculate wtc
        dWTC = CDbl(Format$(App.Frame.Evaluate("atan2(" & Sqr(X * X + Y * Y) & "," & Z & ")"), "0.####"))

        ' ensure positive value
        If (dWAA < 0) Then dWAA = (360 + dWAA)
        If (dWAC < 0) Then dWAC = (360 + dWAC)
        If (dWTC < 0) Then dWTC = (360 + dWTC)

        WAA = dWAA
        WAC = dWAC
        WTC = dWTC

End Sub

Public Sub g_SetNestExtension(ByVal iNestExt As AlphaNestExtension, ByVal bActive As Boolean, _
                              Optional sCustomName As String = vbNullString, Optional lCustomID As Long = 1)
               
        Dim NE                      As Object   ' ACAMNESTLib.NestExtension
        Dim strName                 As String
        Dim lngID                   As Long
        
        '    INTERNAL NAME          SUB ID      DISPLAY NAME
        '    -------------          ------      ------------
        '
        '    Stock...
        '
        '    ToolPathSort           4           Cut Whole Part Together
        '    ToolPathSort           5           Suppress Final Sort
        '    PartAsGroup            1           Group Each Part Separately
        '    RotatedFirst           1           Try Rotated Part First on all Parts
        '    SuppressRedraw         1           Suppress Redraw
        '    PostProcessGroups      1           Remove Groups
        '
        '    Stock, but Have Nestlist Properties...
        '
        '    AssistedNest           1           Assisted Nest
        '    PreserveSheetEdge      1           Leave Edge Gap Uncut
        '    ToolPathSort           1           Minimize Tool Changes
        '    ToolPathSort           2           Drill then Cut Inner Paths First
        '    ToolPathSort           3           Order By Part
        '    RepeatRowColumn        1           Repeat First Row/Column
        '
        '    NestExtensions Add-in...
        '
        '    CutSmallFirst          1           Cut Small Parts First
        '    QuantityMerge          1           Merge Like Part Quantities
        '    QuantityMultiplier     1           Part Quantity Multiplier
        '    OnionSkin              1           Onion Skin Small Parts
        '
        '    NestSmallToLarge Add-In...
        '
        '    NestSmallToLarge       1           Nest Small Parts First
        
        ' set default ID, others will be overwritten
        lngID = 1
        
        Select Case iNestExt
                Case alphaNestExt_CUT_HOLE_PART: strName = "ToolPathSort": lngID = 4
                Case alphaNestExt_SUPPRESS_FINAL_SORT: strName = "ToolPathSort": lngID = 5
                Case alphaNestExt_GROUP_EACH_PART: strName = "PartAsGroup"
                Case alphaNestExt_ROTATED_PARTS_FIRST: strName = "RotatedFirst"
                Case alphaNestExt_SUPPRESS_REDRAW: strName = "SuppressRedraw"
                Case alphaNestExt_REMOVE_GROUPS: strName = "PostProcessGroups"
                Case alphaNestExt_CUT_SMALL_FIRST: strName = "CutSmallFirst"
                Case alphaNestExt_QUANTITY_MERGE: strName = "QuantityMerge"
                Case alphaNestExt_QUANTITY_MULTIPLIER: strName = "QuantityMultiplier"
                Case alphaNestExt_ONION_SKIN: strName = "OnionSkin"
                Case alphaNestExt_NEST_SMALL_TO_LARGE: strName = "NestSmallToLarge"
                Case Else: strName = sCustomName: lngID = lCustomID
        End Select
        
        If gb_HasNestExtension(strName, NE) Then Call NE.SetState(lngID, Abs(bActive))
        
End Sub

Public Function gb_HasNestExtension(ByVal sExtensionName As String, NE As Object, Optional lIndexRet As Long) As Boolean
                                   'ByVal sExtensionName As String, NE As ACAMNESTLib.NestExtension, Optional lIndexRet As Long) As Boolean
                
        Dim NES                     As Object   ' ACAMNESTLib.NestExtensions
        Dim lngIndex                As Long
        Dim lngCount                As Long
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap

        blnRet = False
        
        Set NES = App.Nesting.Extensions
        
        ' validate
        If (NES Is Nothing) Then GoTo Controlled_Exit
        
        lngCount = NES.count
                
        ' lets look for it
        For lngIndex = 1 To lngCount
                
                Set NE = NES(lngIndex)
                
                If Not (NE Is Nothing) Then
                                                                                        
                        ' if we've found what we're looking for, set return val and bail
                        If (StrComp(sExtensionName, NE.name, vbTextCompare) = 0) Then
                                lIndexRet = lngIndex
                                blnRet = True
                                Exit For
                        End If
                        
                End If
                
        Next lngIndex
        
Controlled_Exit:
        
        gb_HasNestExtension = blnRet
                
        Set NES = Nothing

Exit Function

ErrTrap:
        
        'Debug.Print "<" & Err.Number & "> " & Err.Description
        blnRet = False
        Resume Controlled_Exit

End Function

Public Sub g_GetMinMaxOpNumbersForSheet(lMin As Long, lMax As Long, psPathsWithinSheet As Paths)

        Dim Pth                     As Path
        Dim lngMin                  As Long
        Dim lngMax                  As Long
    
On Error GoTo ErrTrap
        
        If (psPathsWithinSheet Is Nothing) Then
                lMin = 0
                lMax = 0
                GoTo Controlled_Exit
        End If
        
        If (psPathsWithinSheet.count = 0) Then
                lMin = 0
                lMax = 0
                GoTo Controlled_Exit
        End If
        
        ' set default min/max operation numbers to the first path's operation number
        lngMin = psPathsWithinSheet(1).OpNo
        lngMax = psPathsWithinSheet(1).OpNo
        
        ' loop thru all paths within the current sheet and find the min/max operation numbers
        For Each Pth In psPathsWithinSheet
            
                With Pth
                        
                        If .IsToolPath Then
                                   
                                If Not .IsPathAllRapids Then
                                   
                                        If (.OpNo > 0) Then
                                                
                                                If (lngMin = 0) Then lngMin = .OpNo
                                                If (lngMax = 0) Then lngMax = .OpNo
                                        
                                                If (.OpNo <= lngMin) Then
                                                        lngMin = .OpNo
                                                ElseIf (.OpNo >= lngMax) Then
                                                        lngMax = .OpNo
                                                End If
                                        End If
                                        
                                End If
                                
                        End If
            
                End With
            
        Next Pth
    
        ' store values to be returned
        lMin = lngMin
        lMax = lngMax
        
        'Debug.Print "    Min Op Number: " & lMin
        'Debug.Print "    Max Op Number: " & lMax
    
Controlled_Exit:
    
        Set Pth = Nothing

Exit Sub
    
ErrTrap:
    
        MsgBox Err.Description, vbExclamation, DEF_APP_TITLE
        Resume Controlled_Exit

End Sub

Public Sub g_TransposeNestedSheetXY(ByVal iQuad As AlphaQuadrant, _
                                    ByVal dQuadX As Double, ByVal dQuadY As Double, _
                                    ByVal dXOriginal As Double, ByVal dYOriginal As Double, _
                                    dXNew As Double, dYNew As Double)
        
        Dim dblRetX                 As Double
        Dim dblRetY                 As Double
        
        ' initialize
        dblRetX = dXOriginal
        dblRetY = dYOriginal

        Select Case iQuad
                
                Case alphaQuadrant_I
                        
                        ' can leave as is
                    
                Case alphaQuadrant_II
                        
                        dblRetX = (dblRetX - dQuadX)
                        
                        ' can leave Y as is
                
                Case alphaQuadrant_III
                
                        dblRetX = (dblRetX - dQuadX)
                        dblRetY = (dblRetY - dQuadY)
                
                Case alphaQuadrant_IV
                
                        ' can leave X as is
                        
                        dblRetY = (dblRetY - dQuadY)
                        
        End Select

        dXNew = dblRetX
        dYNew = dblRetY

End Sub

Public Sub g_LtoG(WP As WorkPlane, dX As Double, dY As Double, dZ As Double)

        Dim dblX                    As Double
        Dim dblY                    As Double
        Dim dblZ                    As Double

        ' if not workplane then return what was given
        If (WP Is Nothing) Then Exit Sub

        dblX = dX
        dblY = dY
        dblZ = dZ
        
        With WP
                dX = .Tmat(0) * dblX + .Tmat(1) * dblY + .Tmat(2) * dblZ + .Tmat(3)
                dY = .Tmat(4) * dblX + .Tmat(5) * dblY + .Tmat(6) * dblZ + .Tmat(7)
                dZ = .Tmat(8) * dblX + .Tmat(9) * dblY + .Tmat(10) * dblZ + .Tmat(11)
        End With
        
End Sub

Public Sub g_GtoL(WP As WorkPlane, dX As Double, dY As Double, dZ As Double)

        Dim dblX                    As Double
        Dim dblY                    As Double
        Dim dblZ                    As Double

        ' if not workplane then return what was given
        If (WP Is Nothing) Then Exit Sub
          
        dblX = dX
        dblY = dY
        dblZ = dZ
        
        With WP
                dX = .Imat(0) * dblX + .Imat(1) * dblY + .Imat(2) * dblZ + .Imat(3)
                dY = .Imat(4) * dblX + .Imat(5) * dblY + .Imat(6) * dblZ + .Imat(7)
                dZ = .Imat(8) * dblX + .Imat(9) * dblY + .Imat(10) * dblZ + .Imat(11)
        End With
        
End Sub

Public Function gps_ConvertTextToGeometry(Optional TS As Texts = Nothing, Optional ByVal bLeaveOriginal As Boolean = False) As Paths
        
        Dim atxText                 As Text
        Dim atxCopy                 As Text
        Dim atxToConvert            As Texts
        Dim atxLine                 As TextLine
        Dim P                       As Path
        Dim PS                      As Paths
        Dim pthsRet                 As Paths
        Dim strText                 As String
        
        If (TS Is Nothing) Then
                Set atxToConvert = App.ActiveDrawing.Text
        Else
                Set atxToConvert = TS
        End If
        
        If (atxToConvert.count > 0) Then
                
                ' create return collection
                Set pthsRet = App.ActiveDrawing.CreatePathCollection
        
                ' convert all text on the active drawing to geometry
                For Each atxText In atxToConvert
                                                                      
                        ' + ADDED a check for text.lines.count as this might be zero if text has
                        '   been edited and the user just backspaces off the text and clicks OK,
                        '   this will leave a text without a line count need to look for dimensions
                        '
                        ' although they are on the DIMENSIONS layer, acam still sees them as text
                        If (UCase$(atxText.GetLayer.name) <> "DIMENSIONS") And (atxText.Lines.count > 0) Then
                                
                                ' leave original?
                                If bLeaveOriginal Then
                                        Set atxCopy = atxText.Copy
                                End If
                                
                                strText = vbNullString
                                
                                For Each atxLine In atxText.Lines
                                        
                                        If (Len(strText) > 0) Then
                                                strText = strText & vbCrLf
                                        End If
                                                                                
                                        strText = strText
                                        
                                Next atxLine
                                                                                
                                Set PS = atxText.ConvertToGeometry
                                
                                If Not (PS Is Nothing) Then

                                        ' tack them on to the return collection
                                        Call g_AppendPathsToCollection(pthsRet, PS)
   
                                End If
                                
                        End If
                        
                Next atxText
                
        End If
        
        Set gps_ConvertTextToGeometry = pthsRet
        
        Set atxText = Nothing
        Set atxLine = Nothing
        Set P = Nothing
        Set PS = Nothing
        Set pthsRet = Nothing
        
End Function

Public Function gb_IsCircle(ByVal P As Path, Optional cpRet As CircleProperties = Nothing) As Boolean
                
        Dim blnRet                  As Boolean

On Error GoTo ErrTrap
                        
        ' try to get circle props
        Set cpRet = P.GetCircleProperties
        
        ' return if circle props exist
        blnRet = Not (cpRet Is Nothing)
                
Controlled_Exit:
        
        ' return
        gb_IsCircle = blnRet

Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gb_IsRectangle(ByVal P As Path, Optional rpRet As RectangleProperties = Nothing) As Boolean
                
        Dim blnRet                  As Boolean

On Error GoTo ErrTrap
                        
        ' try to get circle props
        Set rpRet = P.GetRectangleProperties
        
        ' return if circle props exist
        blnRet = Not (rpRet Is Nothing)
                
Controlled_Exit:
        
        ' return
        gb_IsRectangle = blnRet

Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gb_IsLine(P As Path, Optional iOrientation As AlphaLineOrientation) As Boolean
        
        Dim blnRet                  As Boolean
        
        blnRet = True
                
        ' if more than one element or only element is an arc, then not good
        Select Case True
                Case (P.Elements.count <> 1), P.Elements(1).IsArc: blnRet = False
        End Select
        
        With P.Elements(1)
                Select Case True
                        Case (PSTol(.StartXL) = PSTol(.EndXL)): iOrientation = alphaLine_VERTICAL
                        Case (PSTol(.StartYL) = PSTol(.EndYL)): iOrientation = alphaLine_HORIZONTAL
                        Case Else: iOrientation = alphaLine_ANGLED
                End Select
        End With
        
        gb_IsLine = blnRet
        
End Function

Public Function gb_IsSameWP(WP1 As WorkPlane, WP2 As WorkPlane, Optional ByVal bCompareOrigin As Boolean = False) As Boolean

        Dim xyzWP1                  As WP_XYZ
        Dim xyzWP2                  As WP_XYZ
        Dim blnRet                  As Boolean
                
On Error Resume Next
                
        blnRet = True
                
        ' check that a workplane has actually been sent to the sub
        If (WP1 Is Nothing) Then
                
                If Not (WP2 Is Nothing) Then
                        blnRet = False
                        GoTo Controlled_Exit
                Else
                        GoTo Controlled_Exit
                End If
        
        End If
                
        If (WP2 Is Nothing) Then
                
                If Not (WP1 Is Nothing) Then
                        blnRet = False
                        GoTo Controlled_Exit
                Else
                        GoTo Controlled_Exit
                End If
        
        End If
        
        'Debug.Print WP1.Name & " : " & WP2.Name
        
        With xyzWP1
        
                With .X
                        .X = WP1.Tmat(0)
                        .Y = WP1.Tmat(1)
                        .Z = WP1.Tmat(2)
                End With
                
                With .Y
                        .X = WP1.Tmat(4)
                        .Y = WP1.Tmat(5)
                        .Z = WP1.Tmat(6)
                End With
                
                With .Z
                        .X = WP1.Tmat(8)
                        .Y = WP1.Tmat(9)
                        .Z = WP1.Tmat(10)
                End With
                
                With .Origin
                        .X = WP1.Tmat(3)
                        .Y = WP1.Tmat(7)
                        .Z = WP1.Tmat(11)
                End With
                
        End With

        With xyzWP2

                With .X
                        .X = WP2.Tmat(0)
                        .Y = WP2.Tmat(1)
                        .Z = WP2.Tmat(2)
                End With
                
                With .Y
                        .X = WP2.Tmat(4)
                        .Y = WP2.Tmat(5)
                        .Z = WP2.Tmat(6)
                End With
                
                With .Z
                        .X = WP2.Tmat(8)
                        .Y = WP2.Tmat(9)
                        .Z = WP2.Tmat(10)
                End With
                
                With .Origin
                        .X = WP2.Tmat(3)
                        .Y = WP2.Tmat(7)
                        .Z = WP2.Tmat(11)
                End With
                
        End With
        
        If bCompareOrigin Then
        
                Select Case True
                        Case (xyzWP1.Origin.X <> xyzWP2.Origin.X): blnRet = False
                        Case (xyzWP1.Origin.Y <> xyzWP2.Origin.Y): blnRet = False
                        Case (xyzWP1.Origin.Z <> xyzWP2.Origin.Z): blnRet = False
                End Select
                
        End If
        
        If blnRet Then

                Select Case True
                        
                        Case (xyzWP1.X.X <> xyzWP2.X.X): blnRet = False
                        Case (xyzWP1.X.Y <> xyzWP2.X.Y): blnRet = False
                        Case (xyzWP1.X.Z <> xyzWP2.X.Z): blnRet = False
                    
                        Case (xyzWP1.Y.X <> xyzWP2.Y.X): blnRet = False
                        Case (xyzWP1.Y.Y <> xyzWP2.Y.Y): blnRet = False
                        Case (xyzWP1.Y.Z <> xyzWP2.Y.Z): blnRet = False
                
                        Case (xyzWP1.Z.X <> xyzWP2.Z.X): blnRet = False
                        Case (xyzWP1.Z.Y <> xyzWP2.Z.Y): blnRet = False
                        Case (xyzWP1.Z.Z <> xyzWP2.Z.Z): blnRet = False
                
                End Select
        
        End If
        
Controlled_Exit:
        
        gb_IsSameWP = blnRet

Exit Function

End Function

Public Function gb_IsAcamColor(ByVal lColorRGB As Long) As Boolean
        
        Dim blnRet                  As Boolean
        
        ' assume no
        blnRet = False
        
        Select Case lColorRGB
                
                Case 0, &H800000, &H8000&, _
                     &H808000, &H80&, &H800080, _
                     &H4080&, &HC0C0C0, &H808080, _
                     &HFF0000, &HFF00&, &HFFFF00, &HFF&, _
                     &HFF00FF, &HFFFF&, &HFFFFFF
                     
                        blnRet = True
                        
        End Select
        
        gb_IsAcamColor = blnRet
        
End Function

Public Function gb_HasWorkVolume(Optional pWV As Path = Nothing) As Boolean
        
        Dim P                       As Path
        Dim blnRet                  As Boolean
                
        blnRet = False
        
        ' loop thru all geos and look for work volume
        For Each P In App.ActiveDrawing.Geometries
                
                If P.IsWorkVolume Then
                        Set pWV = P
                        blnRet = True
                        Exit For
                End If
                
        Next P
        
        gb_HasWorkVolume = blnRet
        
        Set P = Nothing
        
End Function

Public Function gb_HasMaterial(Optional pMaterial As Path = Nothing) As Boolean
        
        Dim P                       As Path
        Dim blnRet                  As Boolean
                
        blnRet = False
        
        ' loop thru all geos and look for work volume
        For Each P In App.ActiveDrawing.Geometries
                
                If P.Billet Then
                        Set pMaterial = P
                        blnRet = True
                        Exit For
                End If
                
        Next P
        
        gb_HasMaterial = blnRet
        
        Set P = Nothing
        
End Function

Public Function gi_NestLevel() As Integer       ' ACAMNESTLib.NestLevel
        
        Dim objNest                 As Object   ' ACAMNESTLib.Nesting
        Dim intRet                  As Integer  ' ACAMNESTLib.NestLevel

On Error GoTo ErrTrap
        
        ' nestLevelNONE = 0
        ' nestLevelADVANCED = 1
        ' nestLevelBASIC = 2
        
        intRet = 0      ' nestLevelNONE
        
        ' any nesting?
        Set objNest = App.Nesting
        
        If (objNest Is Nothing) Then GoTo Controlled_Exit
        
        intRet = objNest.Level
        
Controlled_Exit:
        
        gi_NestLevel = intRet
        
        Set objNest = Nothing

Exit Function

ErrTrap:
        
        intRet = 0
        Resume Controlled_Exit

End Function

Public Function gb_HasNesting(oNestInformation As Object) As Boolean    ' NestInformation) As Boolean
        
        Dim blnRet                  As Boolean

On Error Resume Next
        
        blnRet = False
        
        ' any nesting?
        Set oNestInformation = App.ActiveDrawing.GetNestInformation
        
        If (oNestInformation Is Nothing) Then GoTo Controlled_Exit
        If (oNestInformation.Sheets.count = 0) Then GoTo Controlled_Exit
        
        ' must have something
        blnRet = True
        
Controlled_Exit:
        
        gb_HasNesting = blnRet

End Function

Public Function gb_HasSheetGeo() As Boolean
        
        Dim P                       As Path
        Dim PS                      As Paths
        Dim blnRet                  As Boolean
        
        blnRet = False
        
        Set PS = App.ActiveDrawing.Geometries
        
        For Each P In PS
                If P.Sheet Then
                        blnRet = True
                        Exit For
                End If
        Next P
        
        Set P = Nothing
        Set PS = Nothing
        
        gb_HasSheetGeo = blnRet
        
End Function

Public Function gb_IsFeatureAvailable(oSolidFeatures As Object) As Boolean   ' SolidFeatures) As Boolean)
        
        Dim blnRet                  As Boolean
                                        
On Error GoTo ErrTrap
                                        
        blnRet = False
                                        
        Set oSolidFeatures = App.ActiveDrawing.SolidInterface
        
        blnRet = Not (oSolidFeatures Is Nothing)
                                
Controlled_Exit:
        
        gb_IsFeatureAvailable = blnRet

Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit
        
End Function

Public Function gb_HasSolids(oSolidFeatures As Object) As Boolean   ' SolidFeatures) As Boolean
        
        Dim blnRet                  As Boolean
                                        
On Error GoTo ErrTrap
                                        
        blnRet = False
                                        
        Set oSolidFeatures = App.ActiveDrawing.SolidInterface
        
        If gb_IsFeatureAvailable(oSolidFeatures) Then blnRet = CBool(oSolidFeatures.Bodies.count)
                                
Controlled_Exit:
        
        gb_HasSolids = blnRet

Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit
                                        
End Function

Public Function gb_IsSTLAvailable(oSTL As Object) As Boolean
        
        Dim blnRet                  As Boolean
                                        
On Error GoTo ErrTrap
                                        
        blnRet = False
                                        
        Set oSTL = App.ActiveDrawing.STLInterface
        
        blnRet = Not (oSTL Is Nothing)
                                
Controlled_Exit:
        
        gb_IsSTLAvailable = blnRet

Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit
        
End Function

'Public Function gb_HasSTL(SI As STL) As Boolean
'
'        Dim blnRet                  As Boolean
'
'On Error GoTo ErrTrap
'
'        blnRet = False
'
'        Set SI = App.ActiveDrawing.STLInterface
'
'        If Not (SI Is Nothing) Then blnRet = CBool(SI.Parts.count)
'
'Controlled_Exit:
'
'        gb_HasSTL = blnRet
'
'Exit Function
'
'ErrTrap:
'
'        blnRet = False
'        Resume Controlled_Exit
'
'End Function

Public Function gb_HasAnything() As Boolean
        
        Dim Drw                     As Drawing
        Dim blnRet                  As Boolean
        
On Error Resume Next
        
        blnRet = False
        
        Set Drw = App.ActiveDrawing
        
        With Drw
                
                Select Case True
                        
                        Case CBool(.Geometries.count), _
                             CBool(.ToolPaths.count), _
                             CBool(.Text.count), _
                             CBool(.Splines.count), _
                             CBool(.Surfaces.count), _
                             CBool(.MachineComponents.count), _
                             CBool(.Clamps.count), _
                             CBool(.Operations.count), _
                             gb_HasSolids(Nothing) ', _
                             gb_HasSTL(Nothing)
                             
                                blnRet = True
                
                End Select
                
        End With
        
        gb_HasAnything = blnRet
        
        Set Drw = Nothing

End Function

Public Function gb_HasAnythingOnWorkPlane(Wrk As WorkPlane) As Boolean
        
        Dim blnRet                  As Boolean
        
On Error Resume Next
        
        blnRet = False
        
        With Wrk
                
                Select Case True
                        
                        Case CBool(.Geometries.count), _
                             CBool(.ToolPaths.count), _
                             CBool(.Text.count), _
                             CBool(.Splines.count)
                             
                                blnRet = True
                
                End Select
                
        End With
        
        gb_HasAnythingOnWorkPlane = blnRet

End Function

Public Function gb_HasVisibleToolpaths() As Boolean
        
        Dim Op                      As Operation
        Dim SubOp                   As SubOperation
        Dim P                       As Path
        Dim PS                      As Paths
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        blnRet = False
                
        For Each Op In App.ActiveDrawing.Operations
                
                For Each SubOp In Op.SubOperations
                        
                        If Not (SubOp.ToolPaths Is Nothing) Then
                        
                                For Each P In SubOp.ToolPaths
                                        If P.Visible Then
                                                blnRet = True
                                                Exit For
                                        End If
                                Next P
                                
                        End If
                        
                        If blnRet Then Exit For
                
                Next SubOp
                
                If blnRet Then Exit For
        
        Next Op
                        
Controlled_Exit:

        gb_HasVisibleToolpaths = blnRet
        
        Set Op = Nothing
        Set SubOp = Nothing
        Set P = Nothing
        Set PS = Nothing

Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gb_HasSubroutines() As Boolean
        
        Dim P                       As Path
        Dim PS                      As Paths
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap

        blnRet = False
        
        Set PS = App.ActiveDrawing.ToolPaths
        
        If (PS Is Nothing) Then GoTo Controlled_Exit
        If (PS.count = 0) Then GoTo Controlled_Exit
        
        For Each P In PS
                
                ' look for sub routine and bail if we find one
                If (P.SubroutineNumber <> 0) Then
                        blnRet = True
                        Exit For
                End If
                        
        Next P
    
Controlled_Exit:

        gb_HasSubroutines = blnRet
        
Exit Function

ErrTrap:
        
        MsgBox Err.Description, vbExclamation
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gb_SetMaterialFromCopy(P As Path, ByVal dTopZ As Double, ByVal dBottomZ As Double) As Boolean
        
        Dim P2                      As Path
        Dim blnRet                  As Boolean
        
        blnRet = False
        
        If (P Is Nothing) Then GoTo Controlled_Exit
        
        ' create copy of original geo
        Set P2 = P.Copy
                
        If Not (P2 Is Nothing) Then blnRet = P2.SetMaterial(dTopZ, dBottomZ)
            
Controlled_Exit:
            
        gb_SetMaterialFromCopy = blnRet
        
        Set P2 = Nothing
        
End Function

Public Function gb_SetGeometryZLevel(P As Path, ByVal dTop As Double, ByVal dBottom As Double, Optional bSurppressErrorMsg As Boolean = False) As Boolean
        
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        ' assume failure
        blnRet = False
        
        ' if no geos, then bail
        If (P Is Nothing) Then GoTo Controlled_Exit
        
        ' look for invalid vals
        If (dBottom > dTop) Then
                
                If Not bSurppressErrorMsg Then
                        MsgBox "Bottom Z cannot be greater than Top Z.", vbExclamation, "Set Geoemtry Z Levels"
                End If
                
                GoTo Controlled_Exit
                
        End If
                
        ' be sure to ignore tool paths
        If Not P.IsToolPath Then
                
                ' NOTE: attribute values MUST be assigned as Double
                '
                With P
                        .Attribute(LicomUKDMBGeoZLevelTop) = dTop
                        .Attribute(LicomUKDMBGeoZLevelBottom) = dBottom
                        Call .Redraw
                End With
                
        End If
        
        ' should be groovy
        blnRet = True

Controlled_Exit:

        Set P = Nothing
        
        gb_SetGeometryZLevel = blnRet
        
Exit Function

ErrTrap:
        
        blnRet = False
        
        If Not bSurppressErrorMsg Then
                MsgBox Err.Description, vbExclamation, App.name
        End If
        
        Resume Controlled_Exit

End Function

Public Function gb_SetGeometryZLevelsMultiple(PS As Paths, ByVal dTop As Double, ByVal dBottom As Double, Optional bSurppressErrorMsg As Boolean = False) As Boolean
        
        Dim P                       As Path
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        ' assume failure
        blnRet = False
        
        ' if no geos, then bail
        If (PS Is Nothing) Then GoTo Controlled_Exit
        
        ' look for invalid vals
        If (dBottom > dTop) Then
                
                If Not bSurppressErrorMsg Then
                        MsgBox "Bottom Z cannot be greater than Top Z.", vbExclamation, "Set Geoemtry Z Levels"
                End If
                
                GoTo Controlled_Exit
                
        End If
        
        For Each P In PS
                
                ' be sure to ignore tool paths
                If Not P.IsToolPath Then
                        
                        ' NOTE: attribute values MUST be assigned as Double
                        '
                        With P
                                .Attribute(LicomUKDMBGeoZLevelTop) = dTop
                                .Attribute(LicomUKDMBGeoZLevelBottom) = dBottom
                                Call .Redraw
                        End With
                        
                End If
        
        Next P
        
        ' should be groovy
        blnRet = True

Controlled_Exit:

        Set P = Nothing
        
        gb_SetGeometryZLevelsMultiple = blnRet
        
Exit Function

ErrTrap:
        
        blnRet = False
        
        If Not bSurppressErrorMsg Then
                MsgBox Err.Description, vbExclamation, App.name
        End If
        
        Resume Controlled_Exit

End Function

Public Function gb_GetGeometryZLevels(P As Path, dTop As Double, dBottom As Double) As Boolean
        
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        ' assume failure
        blnRet = False
        
        ' if no geos, then bail
        If (P Is Nothing) Then GoTo Controlled_Exit
                
        ' be sure to ignore tool paths
        If Not P.IsToolPath Then
                
                With P
                        
                        If IsEmpty(.Attribute(LicomUKDMBGeoZLevelTop)) Then GoTo Controlled_Exit
                        If IsEmpty(.Attribute(LicomUKDMBGeoZLevelBottom)) Then GoTo Controlled_Exit
                        
                        dTop = .Attribute(LicomUKDMBGeoZLevelTop)
                        dBottom = .Attribute(LicomUKDMBGeoZLevelBottom)
                        
                End With
                
        End If
        
        ' should be groovy
        blnRet = True

Controlled_Exit:
        
        gb_GetGeometryZLevels = blnRet
        
Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gs_GetOperationType(ByVal MD As MillData, Optional sIconString As String = vbNullString) As String
        
        Dim strRet                  As String
        
        Select Case MD.ProcessType2
        
                Case acamProcessBORE
                        strRet = gs_ReadAcamCTX(2332, 10, "Bore Holes")
                        sIconString = "Drill"
                        
                Case acamProcessCONTOUR_POCKET
                        strRet = gs_ReadAcamCTX(2332, 2, "Contour Pocket")
                        sIconString = "Pocket"
                        
                Case acamProcessDRILL
                        strRet = gs_ReadAcamCTX(2332, 7, "Drill Holes")
                        sIconString = "Drill"
                        
                Case acamProcessENGRAVE
                        strRet = gs_ReadAcamCTX(2332, 5, "Engrave")
                        sIconString = "Engrave"
                        
                Case acamProcessLINEAR_POCKET
                        strRet = gs_ReadAcamCTX(2332, 4, "Linear Pocket")
                        sIconString = "Pocket"
                        
                Case acamProcessMACHINE_POLYLINE
                        strRet = gs_ReadAcamCTX(2333, 12, "Along Spline or Polyline")
                        sIconString = "Polyline"
                        
                Case acamProcessMACHINE_SURFACE
                        strRet = gs_ReadAcamCTX(2333, 1, "MC Surfaces")
                        sIconString = "Surface"
                        
                Case acamProcessMANUAL
                        strRet = gs_ReadAcamCTX(2330, 4, "Manual Entry")
                        sIconString = "Manual"
                        
                Case acamProcessPECK
                        strRet = gs_ReadAcamCTX(2332, 8, "Peck Holes")
                        sIconString = "Drill"
                        
                Case acamProcessROUGH_FINISH
                        
                        If (MD.Stock <> 0) Then
                                strRet = gs_ReadAcamCTX(2332, 1, "Roughing Pass")
                        Else
                                strRet = gs_ReadAcamCTX(2330, 7, "Finish Pass")
                        End If
                        
                        sIconString = "RoughFinish"
                        
                Case acamProcessSPIRAL_POCKET
                        strRet = gs_ReadAcamCTX(2332, 3, "Spiral Pocket")
                        sIconString = "Pocket"
                        
                Case acamProcessTAP
                        strRet = gs_ReadAcamCTX(2332, 9, "Tap Holes")
                        sIconString = "Drill"
                
                Case acamProcessCUT_WITH_SAW
                        strRet = gs_ReadAcamCTX(12830, 13, "Sawing")
                        sIconString = "Saw"
                
                Case Else
                        strRet = "Machining"
                        sIconString = "Tool"
                        
        End Select
        
        strRet = UCase$(strRet)
  
        gs_GetOperationType = strRet
  
End Function

Public Function gs_RawModuleName() As String
        
        Dim intLevel                As AlphaProgramLevel
        Dim strRet                  As String
                        
        strRet = vbNullString
        intLevel = gi_ModuleLevel
                
        Select Case gi_ModuleType
                
                Case alphaMod_MILL
                        
                        Select Case intLevel
                                Case alphaLevel_CAD: strRet = "ACADAPS"
                                Case alphaLevel_BASIC: strRet = "BMILLAPS"
                                Case alphaLevel_STANDARD: strRet = "SMILLAPS"
                                Case alphaLevel_ADVANCED: strRet = "AMILLAPS"
                                Case alphaLevel_ADVANCED3D3AXIS: strRet = "AM3AXAPS"
                                Case alphaLevel_ADVANCED3D5AXIS: strRet = "AM5AXAPS"
                                Case alphaLevel_VIEWPLUS: strRet = "VMILLAPS"
                                Case alphaLevel_OEM5AXIS: strRet = "OMILLAPS"
                        End Select
                
                Case alphaMod_ROUTER
                
                        Select Case intLevel
                                Case alphaLevel_BASIC: strRet = "BROUTAPS"
                                Case alphaLevel_STANDARD: strRet = "SROUTAPS"
                                Case alphaLevel_ADVANCED: strRet = "AROUTAPS"
                                Case alphaLevel_ADVANCED3D3AXIS: strRet = "AR3AXAPS"
                                Case alphaLevel_ADVANCED3D5AXIS: strRet = "AR5AXAPS"
                                Case alphaLevel_VIEWPLUS: strRet = "VROUTAPS"
                                Case alphaLevel_PROUTAPS: strRet = "PROUTAPS"
                                Case alphaLevel_OEM5AXIS: strRet = "OROUTAPS"
                        End Select

                Case alphaMod_STONE
                
                        Select Case intLevel
                                Case alphaLevel_BASIC: strRet = "BMARBAPS"
                                Case alphaLevel_STANDARD: strRet = "SMARBAPS"
                                Case alphaLevel_ADVANCED: strRet = "AMARBAPS"
                                Case alphaLevel_ADVANCED3D3AXIS: strRet = "AMAR3APS"
                                Case alphaLevel_ADVANCED3D5AXIS: strRet = "AMAR5APS"
                                Case alphaLevel_VIEWPLUS: strRet = "VMARBAPS"
                                Case alphaLevel_OEM5AXIS: strRet = "OMARBAPS"
                        End Select
                        
                Case alphaMod_LATHE
                        
                        Select Case intLevel
                                Case alphaLevel_BASIC: strRet = "BTURNAPS"
                                Case alphaLevel_STANDARD: strRet = "STURNAPS"
                                Case alphaLevel_ADVANCED: strRet = "ATURNAPS"
                                Case alphaLevel_ADVANCED3D3AXIS: strRet = "AT3AXAPS"
                                Case alphaLevel_ADVANCED3D5AXIS: strRet = "AT5AXAPS"
                                Case alphaLevel_VIEWPLUS: strRet = "VTURNAPS"
                        End Select
                                        
                Case alphaMod_WIRE
                
                        Select Case intLevel
                                Case alphaLevel_STANDARD: strRet = "SWIREAPS"
                                Case alphaLevel_ADVANCED: strRet = "AWIREAPS"
                                Case alphaLevel_VIEWPLUS: strRet = "VWIREAPS"
                        End Select
                        
                Case alphaMod_PROFILING
        
                        Select Case intLevel
                                Case alphaLevel_ADVANCED: strRet = "ALASEAPS"
                                Case alphaLevel_ADVANCED3D5AXIS: strRet = "AL5AXAPS"
                                Case alphaLevel_PROUTAPS: strRet = "VLASEAPS"
                        End Select
        
        End Select
        
        gs_RawModuleName = strRet

End Function

Public Function go_GetSawMillDataFromPath(P As Path) As SAW_MILLDATA_PROPERTIES
        
        Dim uRet                    As SAW_MILLDATA_PROPERTIES
        
        If (P Is Nothing) Then Exit Function
        
        With uRet
                
                ' hidden atts
                .SawInternalCorners = P.Attribute(LicomUKDMBSaw1)
                .SawExternalCorners = P.Attribute(LicomUKDMBSaw2)
                .SawHeadPosition = P.Attribute(LicomUKDMBSaw3)
                .SawInternalCornerDistance = P.Attribute(LicomUKDMBSaw4)
                .SawExternalCornerDistance = P.Attribute(LicomUKDMBSaw5)
                .SawResetAngles = P.Attribute(LicomUKDMBResetAngles)
                .SawIncludeArcs = P.Attribute(LicomUKDMBSawIncludeArcs)
                .SawMinimumArcRadius = P.Attribute(LicomUKDMBSawMinRad)
                
                ' visible atts
                .SawAngle = CDbl(P.Attribute(LicomUKDMBSawAngle))
                
        End With
        
        go_GetSawMillDataFromPath = uRet

End Function

Public Function gb_StartFileNew(Optional bForce As Boolean = False) As Boolean
        
        Dim blnRet                  As Boolean
        Dim strMsgText              As String
        
        blnRet = False
        
        ' if not forcing new drawing then ask the user
        If Not bForce Then
        
                With App.ActiveDrawing
                        If Not gb_HasAnything Then
                                blnRet = True
                                GoTo Controlled_Exit
                        End If
                End With
                
                strMsgText = "This command will clear the current drawing. Any unsaved data will be lost."
                
                If (MsgBox(strMsgText, vbOKCancel) = vbCancel) Then Exit Function
        
        End If
    
        blnRet = True
        
        With App
                .New
                With .ActiveDrawing
                        .ThreeDViews = False
                        Call g_Redraw
                End With
        End With
        
        DoEvents
        
Controlled_Exit:
        
        gb_StartFileNew = blnRet
    
Exit Function
    
End Function

Public Function gb_PostVariableExists(ByVal sVariableName As String) As Boolean
        
        Dim strRet                  As String
        Dim blnRet                  As Boolean

On Error GoTo ErrTrap
        
        ' attempt to get the value of the post variable
        ' if it does not exist, an error is returned
        strRet = App.GetPostUserVariable(sVariableName)
        
        ' if we've made it here, must be OK
        blnRet = True
        
Controlled_Exit:
        
        gb_PostVariableExists = blnRet

ErrTrap:
        
Exit Function
        
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gb_OutputNC(ByVal sOutputFile As String, ByVal iOutputTo As AcamOutNc, _
                            ByVal bVisibleOnly As Boolean, Optional sPost As String = vbNullString) As Boolean
        
        Dim FSO                     As New Scripting.FileSystemObject
        Dim strPostOriginal         As String
        Dim strPost                 As String
        Dim blnRet                  As Boolean
        
On Error Resume Next
                
        blnRet = False
        
        ' store original post
        strPostOriginal = App.PostFileName
        
        ' look for passed post
        If (Len(Trim$(sPost)) = 0) Then
                strPost = App.PostFileName
        Else
                strPost = sPost
        End If
        
        With App
                Call .SelectPost(strPost)
                Call .ActiveDrawing.OutputNC(sOutputFile, iOutputTo, bVisibleOnly)
        End With
        
        DoEvents
        
        Select Case True

                ' look for error, missing file, or file with no size
                Case (Err.Number <> 0)
                        
                        ' not good, clear the error
                        Debug.Print "! Output NC ERROR: " & Err.Description
                        Err.Clear
                        
                Case Not FSO.FileExists(sOutputFile), (FSO.GetFile(sOutputFile).Size = 0)
                        
                        ' nothing created
                        
                Case Else
                        
                        ' i'm alright, don't nobody worry 'bout me
                        blnRet = True
                                                
        End Select
    
Controlled_Exit:
    
        ' reset original post
        Call App.SelectPost(strPostOriginal)
        
        Set FSO = Nothing
        
        gb_OutputNC = blnRet
        
Exit Function
        
End Function

Public Function gb_SaveDrawing(Optional bForce As Boolean = False) As Boolean
 
        Dim strMsgText              As String
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        blnRet = False
        
        ' if already saved then we're OK
        If Not App.ActiveDrawing.Modified Then blnRet = True: GoTo Controlled_Exit
        
        ' if not forcing new drawing then ask the user
        If bForce Then Call App.ActiveDrawing.Save: GoTo Controlled_Exit
                
        strMsgText = "Please ensure the active drawing is saved before you continue. Do you wish to save now?"

        Select Case MsgBox(strMsgText, vbQuestion + vbYesNoCancel)
                Case vbYes: Call App.ActiveDrawing.Save: blnRet = True  ' save and continue
                Case vbNo: blnRet = True                                ' don't save, but continue
                Case vbCancel: GoTo Controlled_Exit                     ' don't save, don't continue
        End Select
        
        ' must be OK
        blnRet = True
        
Controlled_Exit:
        
        gb_SaveDrawing = blnRet

Exit Function
    
ErrTrap:
        
        MsgBox Err.Description, vbExclamation
        blnRet = False
        Resume Controlled_Exit
    
End Function

Public Function gs_LICOMDAT(Optional ByVal bIncludeBackslash As Boolean = True) As String
        
        Dim strRet                  As String
        
        strRet = gs_EnsureBackslash(App.LicomdatPath) & "LICOMDAT"
        
        If bIncludeBackslash Then strRet = gs_EnsureBackslash(strRet)
        
        gs_LICOMDAT = strRet
        
End Function

Public Function gs_LICOMDIR(Optional ByVal bIncludeBackslash As Boolean = True) As String
        
        Dim strRet                  As String
        
        strRet = gs_EnsureBackslash(App.LicomdirPath) & "LICOMDIR"
        
        If bIncludeBackslash Then strRet = gs_EnsureBackslash(strRet)
        
        gs_LICOMDIR = strRet
        
End Function

Public Function gs_TemplateDir(Optional ByVal bIncludeBackslash As Boolean = True) As String
        
        Dim strRet                  As String
        
        strRet = gs_LICOMDIR
        strRet = strRet & "Templates"
        
        If gb_EnsureDirExistance(strRet) Then
                If bIncludeBackslash Then strRet = gs_EnsureBackslash(strRet)
        Else
                strRet = vbNullString
        End If
        
        gs_TemplateDir = strRet
        
End Function

Public Function gs_ToolsDir() As String
        
        Dim strRet                  As String
        
        ' build path to tools folder
        strRet = gs_LICOMDAT
        strRet = strRet & ms_ProgramLetter(True)
        strRet = strRet & "TOOLS.ALP"
        
        gs_ToolsDir = strRet
                        
End Function

Public Function gs_PostDir() As String
        
        Dim strRet                  As String
        
        ' build path to tools folder
        strRet = gs_LICOMDAT
        strRet = strRet & ms_ProgramLetter(True)
        strRet = strRet & "POSTS.ALP"
        
        gs_PostDir = strRet
                        
End Function

Public Function gs_StylesDir(Optional ByVal bIncludeBackslash As Boolean = True) As String
        
        Dim strRet                  As String
        
        strRet = gs_LICOMDIR & "Styles"
        
        If bIncludeBackslash Then strRet = gs_EnsureBackslash(strRet)
        
        gs_StylesDir = strRet
        
End Function

Public Function gb_StyleExists(ByVal sStyle As String, msRet As MillStyle) As Boolean
        
        Dim MS                      As MillStyle
        Dim MSS                     As MillStyles
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        blnRet = False
        
        ' loop thru all available styles and look for the one we want
        For Each MS In App.MillMachiningStyles
                If (StrComp(MS.FileName, sStyle, vbTextCompare) = 0) Then
                        Set msRet = MS
                        blnRet = True
                        Exit For
                End If
        Next MS
        
        ' if we didn't find it, then let user know about it
        If Not blnRet Then
                MsgBox "Unable to find Machining Style.", vbInformation
                Set msRet = Nothing
        End If
        
Controlled_Exit:

        Set MS = Nothing
        Set MSS = Nothing
        
        gb_StyleExists = blnRet

Exit Function

ErrTrap:

        blnRet = False
        Resume Controlled_Exit
        
End Function


Public Function gb_AssignImageToGeometry(P As Path, ByVal sImage As String, Optional ByVal bRedraw As Boolean = True) As Boolean

        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        blnRet = True
        
        P.Attribute(LicomUKDMBBmpName) = sImage
        
        blnRet = CBool(P.GetBitmap.HBITMAP)
        
        If blnRet Then
                If bRedraw Then Call App.ActiveDrawing.Refresh
        End If

Controlled_Exit:

        gb_AssignImageToGeometry = blnRet

Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gb_RemoveImageFromGeometry(P As Path, Optional ByVal bRedraw As Boolean = True) As Boolean

        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        blnRet = True
        
        Call P.DeleteAttribute(LicomUKDMBBmpName)
        
        blnRet = (P.GetBitmap Is Nothing)
        
        If blnRet Then
                If bRedraw Then Call App.ActiveDrawing.Refresh
        End If

Controlled_Exit:

        gb_RemoveImageFromGeometry = blnRet

Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gb_IsAlphacamLayer(ByVal Lyr As Layer) As Boolean
        
        Dim blnRet                  As Boolean
        
        blnRet = False
        
        If (Lyr Is Nothing) Then GoTo Controlled_Exit
        
        Select Case True
                
                Case (Lyr.Special <> 0), gb_IsMachineOrClampLayer(Lyr)
                        
                        blnRet = True
        
        End Select
        
Controlled_Exit:

        gb_IsAlphacamLayer = blnRet

Exit Function
        
End Function

Public Function gb_IsMachineOrClampLayer(ByVal Lyr As Layer) As Boolean
        
        Dim blnRet                  As Boolean
        
        blnRet = False
        
        If Not (Lyr Is Nothing) Then
                
                Select Case True
                        
                        Case CBool(Lyr.Attribute(LicomUKDMBMCType)), _
                             CBool(Lyr.Attribute(LicomUKSAJFixtureLayer))
                             
                                blnRet = True
                
                End Select
                
        End If
        
Controlled_Exit:
        
        gb_IsMachineOrClampLayer = blnRet
        
Exit Function

End Function

Public Function gb_IsThisActivePost() As Boolean

        Dim blnRet                  As Boolean

On Error Resume Next
        
        blnRet = (StrComp(gs_ThisFile, App.PostFileName, vbTextCompare) = 0)
        
        gb_IsThisActivePost = blnRet

End Function

Public Function gb_IsDrawingSaved(Optional sFullNameRet As String = vbNullString) As Boolean
        
        Dim strMsg                  As String
        Dim intRet                  As VbMsgBoxResult
        Dim blnRet                  As Boolean
                        
        blnRet = False
                
        Select Case True
        
                Case (App.ChangeNumber = 0): blnRet = True
                                                    
                ' ok, stick with me now...if we ChangeNumber <> 0, we
                ' could have a scenario where the user has created
                ' some user layers, but then deleted everything on
                ' them. in this instance, acam will not allow saving
                ' of the file.  so, we need to simply flag it as
                ' though it has already been saved so that the user
                ' doesn't get stuck chasing their tail.
                Case Not gb_HasAnything: blnRet = True
                
                Case Else
    
                        strMsg = gs_ReadAcamCTX(2298, 1, "Drawing Changed but Not Saved") & "."
                        strMsg = strMsg & vbCrLf
                        strMsg = strMsg & gs_ReadAcamCTX(2298, 2, "Do you want to Save it Now?")
                        
                        intRet = MsgBox(strMsg, vbQuestion + vbYesNoCancel, App.name)
                        
                        Select Case intRet
                                
                                Case vbYes
                                        
                                        ' save drawing
                                        Call App.ActiveDrawing.Save
                                        
                                        ' if we still have changes, then bail
                                        blnRet = (App.ChangeNumber = 0)
                                        
                                Case vbNo: blnRet = True ' force to true as we'll ignore changes
                                Case vbCancel   ' already false
                                
                        End Select
                                                        
        End Select
                        
        If blnRet Then sFullNameRet = App.ActiveDrawing.FullName
                        
        gb_IsDrawingSaved = blnRet

End Function

Public Function gb_PathsAreOnSamePlane(P1 As Path, P2 As Path) As Boolean

        Dim intThisFace             As AlphaWorkFace
        Dim intNextFace             As AlphaWorkFace
        Dim blnRet                  As Boolean
        
On Error Resume Next
        
        ' default to true
        blnRet = False
                
        If (P1 Is Nothing) Then GoTo Controlled_Exit
        If (P2 Is Nothing) Then GoTo Controlled_Exit
        
        ' get this and next face
        intThisFace = gi_GetWVF(P1.GetWorkPlane)
        intNextFace = gi_GetWVF(P2.GetWorkPlane)
        
        blnRet = (intThisFace = intNextFace)
                        
Controlled_Exit:
        
        gb_PathsAreOnSamePlane = blnRet

End Function

'Public Function gb_PickSTL(ByVal sPrompt As String, Optional oSelectedParts As Collection = Nothing) As Boolean
'
'        Dim oSTL                    As STL
'        Dim oPart                   As stlPart
'        Dim blnRet                  As Boolean
'
'        blnRet = False
'
'        If Not gb_HasSTL(oSTL) Then GoTo Controlled_Exit
'
'        Set oSelectedParts = New Collection
'
'        With App.ActiveDrawing
'
'                Call .SetGeosSelected(False)
'                Call .SetToolPathsSelected(False)
'
'                blnRet = .UserSelectMultiAddinObjects2(sPrompt, acamSelectDRAW_SELECTED, "inputstl")
'
'                If blnRet Then
'                        For Each oPart In oSTL.Parts
'                                If oPart.Selected Then Call oSelectedParts.add(oPart)
'                        Next oPart
'                End If
'
'        End With
'
'Controlled_Exit:
'
'        Set oPart = Nothing
'        Set oSTL = Nothing
'
'        gb_PickSTL = blnRet
'
'Exit Function
'
'End Function

Public Function gb_PickTool(Optional sTool As String = "$USER", Optional MT As MillTool = Nothing) As Boolean

On Error Resume Next
        
        Set MT = App.SelectTool(sTool)
                
        gb_PickTool = Not (MT Is Nothing)
    
Controlled_Exit:
    
Exit Function

End Function

Public Function gb_OpenTool(ByVal sTool As String, Optional MT As MillTool = Nothing) As Boolean

On Error Resume Next
        
        Set MT = App.OpenTool(sTool)
                
        gb_OpenTool = Not (MT Is Nothing)
    
Controlled_Exit:
    
Exit Function

End Function

Public Function gs_AcamFileType(ByVal iType As AlphaFileType) As String

        Dim strRet                  As String
        Dim strChrMod               As String
        Dim strType                 As String
        Dim strMod                  As String
        Dim strDefault              As String
        
        ' always start with this...
        strRet = "Alphacam" & " "

        ' !! not a standard file type !!
        If (iType = alphaFile_TEMPLATE) Then
                strRet = strRet & "Drawing Template"
                GoTo Controlled_Exit
        End If

        ' now get the program (module) letter and name
        strChrMod = ms_ProgramLetter(False, strMod)
        
        ' append the mod name
        strRet = strRet & strMod & " "
        
        ' get the default name
        Select Case iType
                Case alphaFile_DRAWING: strDefault = "Drawing"
                Case alphaFile_NC: strDefault = "NC Program"
                Case alphaFile_FONT: strDefault = "User Font"
                Case alphaFile_POST: strDefault = "Post"
                Case alphaFile_TOOL: strDefault = "Tool"
                Case alphaFile_VBAPROJECT: strDefault = "VBA Project"
                Case alphaFile_STYLE: strDefault = "Style"
                Case alphaFile_MACHINECONFIG: strDefault = "Machine"
                Case alphaFile_THREAD: strDefault = "Thread"
                Case alphaFile_MACRO: strDefault = "Macro"
                Case alphaFile_NESTLIST: strDefault = "Nest List"
        End Select
        
        ' and finally, get the actual name
        Select Case iType
                Case alphaFile_VBAPROJECT: strType = gs_ReadAeditCTX(1932, 1, strDefault)
                Case alphaFile_STYLE: strType = gs_ReadAeditCTX(1933, 1, strDefault)
                Case alphaFile_MACHINECONFIG: strType = gs_ReadAeditCTX(1934, 1, strDefault)
                Case Else: strType = gs_ReadAeditCTX(1930, iType, strDefault)
        End Select
        
        strRet = strRet & strType
        
Controlled_Exit:
        
        gs_AcamFileType = strRet

End Function

Public Function gs_AcamExt(ByVal iType As AlphaExtensionType, ByVal bIncludePoint As Boolean, _
                           Optional ByVal bUcase As Boolean = False) As String
        
        Dim strRet                  As String
        Dim strChrMod               As String
        Dim strChrType              As String
                
        Select Case iType
                
                Case alphaExtension_NC: strRet = "anc"
                Case alphaExtension_NESTLIST: strRet = "anl"
                Case alphaExtension_DXF: strRet = "dxf"
                Case alphaExtension_DWG: strRet = "dwg"
                Case alphaExtension_IGES: strRet = "iges"
                Case alphaExtension_EMF: strRet = "emf"
                Case alphaExtension_WMF: strRet = "wmf"
                Case alphaExtension_DAT, alphaExtension_MATERIAL: strRet = "dat"
                Case alphaExtension_TEXT: strRet = "txt"
                Case alphaExtension_VBAPOSTCONFIG: strRet = "apc"
                Case alphaExtension_TEMPLATE: strRet = "adt"
                Case Else
                
                        Select Case iType
                                Case alphaExtension_DRAWING: strChrType = "d"
                                Case alphaExtension_FONT: strChrType = "f"
                                Case alphaExtension_POST: strChrType = "p"
                                Case alphaExtension_TOOL: strChrType = "t"
                                Case alphaExtension_VBA: strChrType = "b"
                                Case alphaExtension_STYLE: strChrType = "y"
                                Case alphaExtension_MACHINECONFIG: strChrType = "mc"
                        End Select
                        
                        strChrMod = ms_ProgramLetter

                        strRet = "a" & strChrMod & strChrType
                
        End Select

        If bIncludePoint Then strRet = "." & strRet
        If bUcase Then strRet = UCase$(strRet)

        gs_AcamExt = strRet

End Function

Public Function gi_ModuleType() As AlphaModuleType
        gi_ModuleType = App.ProgramLetter
End Function

Public Function gi_ModuleLevel() As AlphaProgramLevel
        gi_ModuleLevel = App.ProgramLevel
End Function

Public Function gi_GetWVF(WP As WorkPlane) As AlphaWorkFace

        '   WVF = ' Return value of Post Variable WVF (Work Volume Face)
        '
        '   0 if Work Plane is TOP face ie seen when looking in Global Z- direction
        '   1 if Work Plane is FRONT face. ie seen when looking in Global Y+ direction
        '   2 if Work Plane is RIGHT face. ie seen when looking in Global X- direction
        '   3 if Work Plane is BACK face. ie seen when looking in Global Y- direction
        '   4 if Work Plane is LEFT face. ie seen when looking in Global X+ direction
        '   5 if Work Plane is BOTTOM face. ie seen when looking in Global Z+ direction.
        '
        '   -1 if face is NOT one of the above.
        '
        '   It is important that the LOCAL Z axis points OUT from the Work Volume on each face.
        
        Dim intRet                  As AlphaWorkFace
        
        Const Tol                   As Double = 0.000001

        Select Case True
                Case (WP Is Nothing): intRet = alphaWorkFace_TOP                       ' 0
                Case (Abs(WP.Tmat(10) - 1) < Tol): intRet = alphaWorkFace_TOP          ' 0
                Case (Abs(WP.Tmat(6) + 1) < Tol): intRet = alphaWorkFace_FRONT         ' 1
                Case (Abs(WP.Tmat(2) - 1) < Tol): intRet = alphaWorkFace_RIGHT         ' 2
                Case (Abs(WP.Tmat(6) - 1) < Tol): intRet = alphaWorkFace_BACK          ' 3
                Case (Abs(WP.Tmat(2) + 1) < Tol): intRet = alphaWorkFace_LEFT          ' 4
                Case (Abs(WP.Tmat(10) + 1) < Tol): intRet = alphaWorkFace_BOTTOM       ' 5
                Case Else: intRet = alphaWorkFace_UNKNOWN                              ' -1
        End Select
        
        gi_GetWVF = intRet
        
End Function

Public Function gs_ClampNameFromID(ByVal lClampID As Long) As String
        
        Dim C                       As Clamp
        Dim CS                      As Clamps
        Dim strRet                  As String
        
        Set CS = App.ActiveDrawing.Clamps
        
        For Each C In CS
        
                If (C.Number = lClampID) Then
                        strRet = C.name
                        Exit For
                End If
                
        Next C
        
        Set C = Nothing
        Set CS = Nothing
        
        gs_ClampNameFromID = strRet

End Function

Public Function gs_Face(ByVal eFace As AlphaWorkFace) As String

        Select Case eFace
                Case alphaWorkFace_TOP: gs_Face = "Top (-Z)"
                Case alphaWorkFace_FRONT: gs_Face = "Front (+Y)"
                Case alphaWorkFace_RIGHT: gs_Face = "Right (-X)"
                Case alphaWorkFace_BACK: gs_Face = "Back (-Y)"
                Case alphaWorkFace_LEFT: gs_Face = "Left (+X)"
                Case alphaWorkFace_BOTTOM: gs_Face = "Bottom (+Z)"
                Case alphaWorkFace_UNKNOWN: gs_Face = "<UNKNOWN>"
        End Select

End Function

'Public Function gwp_CreateWP(ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double, _
'                             ByVal dRotation As Double, ByVal dTilt As Double) As WorkPlane
'
'        Dim wrkRet                  As WorkPlane
'
'On Error GoTo ErrTrap
'
'        Set wrkRet = App.ActiveDrawing.CreateWorkPlane(dX, dY, dZ, _
'                                     dX + gd_CosDeg(dRotation + 90), dY + gd_SinDeg(dRotation + 90), dZ, _
'                                     dX - gd_CosDeg(dRotation) * gd_CosDeg(dTilt), _
'                                     dY - gd_SinDeg(dRotation) * gd_CosDeg(dTilt), _
'                                     dZ + gd_SinDeg(dTilt))
'
'Controlled_Exit:
'
'        Set gwp_CreateWP = wrkRet
'        Set wrkRet = Nothing
'
'Exit Function
'
'ErrTrap:
'
'        Set wrkRet = Nothing
'        Resume Controlled_Exit
'
'End Function

Public Sub g_ShowAllUserLayers(ByVal bIncludeMachineConfig As Boolean, ByVal bIncludeClamps As Boolean)
        
        Dim Lyr                     As Layer
        
        For Each Lyr In App.ActiveDrawing.Layers
                
                If Not CBool(Lyr.Special) Then
                        
                        Select Case True
                                
                                Case (Lyr.Attribute(LicomUKDMBMCType) <> 0)
                                        
                                        If bIncludeMachineConfig Then Lyr.Visible = True
                                
                                Case (Lyr.Attribute(LicomUKSAJFixtureLayer) <> 0)
                                
                                        If bIncludeClamps Then Lyr.Visible = True
                                
                                Case Else: Lyr.Visible = True
                                
                        End Select
                        
                End If
                
        Next Lyr
        
        Set Lyr = Nothing
        
End Sub

Public Sub g_HideAllUserLayers(ByVal bIncludeMachineConfig As Boolean, ByVal bIncludeClamps As Boolean)
        
        Dim Lyr                     As Layer
        
        For Each Lyr In App.ActiveDrawing.Layers
        
                If Not CBool(Lyr.Special) Then

                        Select Case True
                                
                                Case (Lyr.Attribute(LicomUKDMBMCType) <> 0)
                                        
                                        If bIncludeMachineConfig Then Lyr.Visible = False
                                
                                Case (Lyr.Attribute(LicomUKSAJFixtureLayer) <> 0)
                                
                                        If bIncludeClamps Then Lyr.Visible = False
                                
                                Case Else: Lyr.Visible = False
                                
                        End Select
                        
                End If
                
        Next Lyr
        
        Set Lyr = Nothing
        
End Sub

Public Function glyr_GetActiveLayer() As Layer

        Dim lyrRet                  As Layer
        Dim Lyr                     As Layer
        Dim Lyrs                    As Layers

On Error Resume Next

        Set Lyrs = App.ActiveDrawing.Layers

        For Each Lyr In Lyrs
                If Lyr.Active Then
                        Set lyrRet = Lyr
                        Exit For
                End If
        Next Lyr

        ' return it
        Set glyr_GetActiveLayer = lyrRet

End Function

Public Function glyr_GetLayer(ByVal sLayerName As String, ByVal bCreate As Boolean, bWasCreated As Boolean) As Layer
        
        Dim lyrLabel                As Layer
        
On Error Resume Next
        
        ' try to get the layer
        Set lyrLabel = App.ActiveDrawing.Layers(sLayerName)
        
        bWasCreated = False
        
        ' if not there then create it
        If (lyrLabel Is Nothing) Then
                If bCreate Then
                        Set lyrLabel = App.ActiveDrawing.CreateLayer(sLayerName)
                        bWasCreated = Not (lyrLabel Is Nothing)
                End If
        End If
                        
        ' return it
        Set glyr_GetLayer = lyrLabel
                
End Function

Public Sub g_HideAllToolpaths(Optional ByVal bRedraw As Boolean = True)
        
        Dim strTP                   As String

        ' get the name of the toolpath layer from the acam language file
        strTP = gs_ReadAcamCTX(2353, 3, "TOOLPATHS")
        
        ' hide all toolpaths
        App.ActiveDrawing.Layers(strTP).Visible = False
        
        If bRedraw Then Call g_Redraw
                
End Sub

Public Sub g_ShowAllToolpaths(Optional ByVal bRedraw As Boolean = True)
        
        Dim strTP                   As String

        ' get the name of the toolpath layer from the acam language file
        strTP = gs_ReadAcamCTX(2353, 3, "TOOLPATHS")
        
        ' hide all toolpaths
        App.ActiveDrawing.Layers(strTP).Visible = True
        
        If bRedraw Then Call g_Redraw
                
End Sub

Public Sub g_HideGeos(Optional psGeos As Paths = Nothing, Optional bSelectedOnly As Boolean = False, Optional ByVal bRedraw As Boolean = True)
        
        Dim P                       As Path
        Dim PS                      As Paths
        
        If (psGeos Is Nothing) Then
                Set PS = App.ActiveDrawing.Geometries
        Else
                Set PS = psGeos
        End If
                
        For Each P In PS
                
                If bSelectedOnly Then
                        If P.Selected Then P.Visible = False
                Else
                        P.Visible = False
                End If
                
        Next P
                        
        If bRedraw Then Call g_Redraw
                        
        Set PS = Nothing
        Set P = Nothing
                
End Sub

Public Sub g_UnhideGeos(Optional psGeos As Paths = Nothing, Optional bSelectedOnly As Boolean = False, Optional ByVal bRedraw As Boolean = True)
        
        Dim P                       As Path
        Dim PS                      As Paths
        
        If (psGeos Is Nothing) Then
                Set PS = App.ActiveDrawing.Geometries
        Else
                Set PS = psGeos
        End If
                
        For Each P In PS
                
                If bSelectedOnly Then
                        If P.Selected Then P.Visible = True
                Else
                        P.Visible = True
                End If
                
        Next P
                        
        If bRedraw Then Call g_Redraw
                        
        Set PS = Nothing
        Set P = Nothing
                
End Sub

Public Sub g_ShowAllGeos(Optional ByVal bRedraw As Boolean = True)

        Dim P                       As Path
        Dim PS                      As Paths
        
        Set PS = App.ActiveDrawing.Geometries
                
        For Each P In PS
                With P
                        If .GetLayer.Visible Then .Visible = True
                End With
        Next P
                        
        If bRedraw Then Call g_Redraw
                        
        Set PS = Nothing
        Set P = Nothing
        
End Sub

Public Sub g_ShowAllOpGeos(Optional ByVal bRedraw As Boolean = True)
        
        Dim P                       As Path
        Dim Op                      As Operation
        Dim Ops                     As Operations
        Dim SubOp                   As SubOperation
        
        ' if no ops then bail
        If (App.ActiveDrawing.Operations.count = 0) Then Exit Sub
        
        Set Ops = App.ActiveDrawing.Operations
        
        For Each Op In Ops
                
                For Each SubOp In Op.SubOperations
                        
                        If SubOp.Visible Then
                                For Each P In SubOp.Geometries
                                        P.Visible = True
                                Next P
                        End If
        
                Next SubOp
        
        Next Op
        
        If bRedraw Then Call g_Redraw
        
        Set P = Nothing
        Set Op = Nothing
        Set Ops = Nothing
        Set SubOp = Nothing
        
End Sub

Public Sub g_WipeSolids(Optional ByVal bRedraw As Boolean = True)

        Dim SF                      As Object   ' SolidFeatures
        Dim SB                      As Object   ' SolidBody

On Error Resume Next
        
        ' 08/17/07 - rg
        '   + MODIFIED for V7.5
        '
        'Set SF = New SolidFeatures
        Set SF = App.ActiveDrawing.SolidInterface
        
        For Each SB In SF.Bodies
                Call SB.Delete
        Next SB
        
        If bRedraw Then Call g_Redraw
        
        Set SB = Nothing
        Set SF = Nothing
        
        If (Err.Number <> 0) Then Call Err.Clear
        
End Sub

Public Sub g_WipeNestList(N As Nesting, ByVal sName As String)

        Dim ANL                     As Object   ' ACAMNESTLib.Nestlist
        Dim lngAll                  As Long
        Dim lngIndex                As Long
                        
On Error Resume Next
        
        If (N Is Nothing) Then Exit Sub
        
        ' get nest list count
        lngAll = N.count
                
        If (lngAll = 0) Then GoTo Controlled_Exit
                
        For lngIndex = 1 To lngAll
                
                Set ANL = N(lngIndex)
                
                ' compare name and wipe nest if match is found
                ' notice that we don't exit the loop as there
                ' could be more than one
                If (StrComp(ANL.ListName, sName, vbTextCompare) = 0) Then
                        Call N.DeleteNestList(ANL.FileName)
                        N.GetNestInformation.Refresh
                        DoEvents
                End If
                        
        Next lngIndex
        
Controlled_Exit:

        If (Err.Number <> 0) Then Err.Clear
        
        Set ANL = Nothing
        
Exit Sub
        
End Sub

Public Sub g_WipeNestAttributesFromPath(P As Path)
    
On Error Resume Next
        
        With P
                Call .DeleteAttribute(LicomUKsab_nest)
                Call .DeleteAttribute(LicomUKsab_path)
                Call .DeleteAttribute(LicomUKsab_part)
                Call .DeleteAttribute(LicomUKsab_outer)
                Call .DeleteAttribute(LicomUKsab_sheet)
                Call .DeleteAttribute(LicomUKja_part)
                Call .DeleteAttribute(LicomUKjba_part)
        End With
        
End Sub

Public Function gps_GetAllPaths() As Paths

        Dim P                       As Path
        Dim PS                      As Paths

On Error GoTo ErrTrap
        
        With App.ActiveDrawing
        
                ' store all geos and toolpaths in a single collection
                Set PS = .CreatePathCollection
                
                If (.Geometries.count > 0) Then

                        For Each P In .Geometries

                                ' look for dimensions/construction
                                If (P.GetLayer.Special <> 0) Then
                                                
                                        Select Case P.GetLayer.name
                                                Case "DIMENSIONS", "CONSTRUCTION" ' do nothing
                                                Case Else: Call PS.add(P)
                                        End Select
                                        
                                Else
                                        Call PS.add(P)
                                End If
                                
                        Next P
                        
                End If
                        
                If (.ToolPaths.count > 0) Then
                        For Each P In .ToolPaths
                                Call PS.add(P)
                        Next P
                End If
                
        End With
        
Controlled_Exit:

        Set gps_GetAllPaths = PS
        
        Set P = Nothing
        Set PS = Nothing
        
Exit Function

ErrTrap:
        
        Set PS = Nothing
        Resume Controlled_Exit
        
End Function

Public Function gps_GetPathsInGroup(ByVal iGroup As Integer, PS As Paths) As Paths

        Dim P                       As Path
        Dim pthsRet                 As Paths
    
On Error GoTo ErrTrap
    
        Set pthsRet = App.ActiveDrawing.CreatePathCollection
    
        For Each P In PS
                If (P.Group = iGroup) Then Call pthsRet.add(P)
        Next P
    
Controlled_Exit:
    
        Set gps_GetPathsInGroup = pthsRet
    
        Set P = Nothing
        Set pthsRet = Nothing
    
Exit Function

ErrTrap:
    
        Set pthsRet = Nothing
        Resume Controlled_Exit
    
End Function

Public Function gps_GetPathsNotInGroup(ByVal iGroup As Integer, PS As Paths) As Paths

        Dim P                       As Path
        Dim pthsRet                 As Paths
    
On Error GoTo ErrTrap
    
        Set pthsRet = App.ActiveDrawing.CreatePathCollection
    
        For Each P In PS
                If (P.Group <> iGroup) Then Call pthsRet.add(P)
        Next P
    
Controlled_Exit:
    
        Set gps_GetPathsNotInGroup = pthsRet
    
        Set P = Nothing
        Set pthsRet = Nothing
    
Exit Function

ErrTrap:
    
        Set pthsRet = Nothing
        Resume Controlled_Exit
    
End Function

Public Function gps_GetToolPathsInOperation(ByVal Op As Operation) As Paths
        
        Dim SubOp                   As SubOperation
        Dim SubOps                  As SubOperations
        Dim P                       As Path
        Dim PS                      As Paths
        Dim pthsRet                 As Paths
        
On Error GoTo ErrTrap
        
        If (Op Is Nothing) Then GoTo Controlled_Exit
        
        Set pthsRet = App.ActiveDrawing.CreatePathCollection
        Set SubOps = Op.SubOperations
        
        ' loop thru all suboperatations within this operation
        ' and get all of their toolpaths
        For Each SubOp In SubOps
                
                Set PS = SubOp.ToolPaths
                
                For Each P In PS
                        Call pthsRet.add(P)
                Next P
                
        Next SubOp
        
Controlled_Exit:
        
        Set gps_GetToolPathsInOperation = pthsRet
        
        Set SubOp = Nothing
        Set SubOps = Nothing
        Set P = Nothing
        Set PS = Nothing
        Set pthsRet = Nothing
        
Exit Function

ErrTrap:
        
        Set pthsRet = Nothing
        Resume Controlled_Exit
        
End Function

Public Sub g_RemovePathFromColleciton(psCollection As Paths, pToRemove As Path, Optional bExisted As Boolean = True)

On Error GoTo ErrTrap

        ' remove - will raise error if not in collection
        Call psCollection.Remove(pToRemove)

Controlled_Exit:

Exit Sub

ErrTrap:
        
        ' if error raised, the most likely not in collection
        bExisted = False
        Resume Controlled_Exit

End Sub

Public Function gb_AttrExistsInPost(ByVal sAttribute As String, lNumber As Long) As Boolean
        
        Dim lngIndex                As Long
        Dim lngCount                As Long
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
    
        blnRet = False
        
        ' get count of existing attributes within the post
        lngCount = App.GetPostAttributeCount
        
        If (lngCount = 0) Then Exit Function
        
        ' look for the attribute we need
        For lngIndex = 1 To lngCount
                If (UCase$(App.GetPostAttributeName(lngIndex)) = UCase$(sAttribute)) Then
                        lNumber = App.GetPostAttributeNumber(lngIndex)
                        blnRet = True
                        Exit For
                End If
        Next lngIndex
        
Controlled_Exit:

        gb_AttrExistsInPost = blnRet
        
Exit Function

ErrTrap:
    
        blnRet = False
        Resume Controlled_Exit
    
End Function

Public Function gb_AttrExists(ByVal sAttr As String, oAcamObject As Object, Optional sAttrVal As String = vbNullString) As Boolean
        
        Dim lngCount                As Long
        Dim lngIndex                As Long
        Dim strName                 As String
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap

        blnRet = False
        
        ' get attribute count of passed alphacam object
        lngCount = oAcamObject.GetAttributeCount
        
        If (lngCount > 0) Then
                
                ' loop thru and look case sensitive match
                For lngIndex = 1 To lngCount
                        
                        strName = oAcamObject.GetAttributeName(lngIndex)
                        
                        If (StrComp(sAttr, strName, vbBinaryCompare) = 0) Then
                                sAttrVal = oAcamObject.Attribute(sAttr)
                                blnRet = True
                                Exit For
                        End If
                        
                Next lngIndex
                
        End If

Controlled_Exit:
        
        gb_AttrExists = blnRet

Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit
        
End Function

Public Function gb_IsMultidrillOp(ByVal Op As Operation) As Boolean
        
        Dim P                       As Path
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        ' will force error if doesn't exists - errtrap will catch it
        Set P = Op.SubOperations(1).ToolPaths(1)
        
        ' return whether or not the path has multidrill attribute
        blnRet = CBool(P.Attribute("AcamUSrg_Mdrill"))
        
Controlled_Exit:

        gb_IsMultidrillOp = blnRet
        
        Set P = Nothing
        
Exit Function

ErrTrap:
        
        blnRet = False
        Resume Controlled_Exit

End Function

'Public Function gp_PathClosestToZero(psTP As Paths) As Path
'
'        Dim Drw                     As Drawing
'        Dim P                       As Path
'        Dim pthRet                  As Path
'        Dim udtXYZ                  As POINT_XYZ
'        Dim udtMinXYZ               As POINT_XYZ
'        Dim dblDistance             As Double
'        Dim dblMinDistance          As Double
'
'On Error GoTo ErrTrap
'
'        If (psTP Is Nothing) Then GoTo Controlled_Exit
'
'        Set Drw = App.ActiveDrawing
'
'         ' init
'         dblMinDistance = 999999999
'
'         udtMinXYZ.X = 0
'         udtMinXYZ.Y = 0
'
'         For Each P In psTP
'
'                Call P.DeleteRapids(True, True)
'
'                ' store the values of the first hole
'                With P.Elements(1)
'                        udtXYZ.X = .StartXG
'                        udtXYZ.Y = .StartYG
'                End With
'
'                dblDistance = gd_Distance(udtMinXYZ, udtXYZ)
'
'                If (dblDistance <= dblMinDistance) Then
'                        dblMinDistance = dblDistance
'                        'udtMinXYZ = udtXYZ
'                        Set pthRet = P
'                End If
'
'         Next P
'
'Controlled_Exit:
'
'        Set gp_PathClosestToZero = pthRet
'
'        Set P = Nothing
'        Set pthRet = Nothing
'        Set Drw = Nothing
'
'Exit Function
'
'ErrTrap:
'
'        MsgBox Err.Description, vbExclamation
'        Set pthRet = Nothing
'        Resume Controlled_Exit
'
'End Function

'Public Function gb_Trace3DPolyline(PL As Path, pRet As Path) As Boolean
'
'        Dim geoLine                 As Geo2D
'        Dim E                       As Element
'        Dim dblX                    As Double
'        Dim dblY                    As Double
'        Dim blnFirst                As Boolean
'        Dim blnRet                  As Boolean
'
'On Error GoTo ErrTrap
'
'        blnRet = True
'        blnFirst = True
'
'        If (PL Is Nothing) Then GoTo Controlled_Exit
'        If Not PL.Is3D Then GoTo Controlled_Exit
'
'        For Each E In PL.Elements
'
'                If blnFirst Then
'                        Set geoLine = App.ActiveDrawing.Create2DGeometry(E.StartXG, E.StartYG)
'                        blnFirst = False
'                End If
'
'                dblX = E.EndXG
'                dblY = E.EndYG
'
'                If Not gb_Equal(dblX, dblY, 0.0001) Then
'                        Call geoLine.AddLine(dblX, dblY)
'                End If
'
'        Next E
'
'        Set pRet = geoLine.Finish
'
'        If (pRet Is Nothing) Then
'                blnRet = False
'        Else
'                Call pRet.SetLayer(PL.GetLayer)
'        End If
'
'Controlled_Exit:
'
'        ' wipe return polyline if if failed to finish
'        If Not blnRet Then
'                On Error Resume Next
'                If Not (pRet Is Nothing) Then pRet.Delete
'        End If
'
'        Set geoLine = Nothing
'        Set E = Nothing
'
'        gb_Trace3DPolyline = blnRet
'
'Exit Function
'
'ErrTrap:
'
'        blnRet = False
'        Resume Controlled_Exit
'
'End Function

Public Function gp_GetWV() As Path

        Dim P                       As Path
        
On Error Resume Next
        
        For Each P In App.ActiveDrawing.Geometries
                If P.IsWorkVolume Then Exit For
        Next P
        
        Set gp_GetWV = P

End Function

Public Function gp_Offset(P As Path, dDistance As Double, iSide As AcamToolSide, Optional bDeleteOriginal As Boolean = False) As Path

        Dim pthsRet                 As Paths

        Set pthsRet = P.Offset(dDistance, iSide)

        ' select and delete original?
        If bDeleteOriginal Then
                Call P.Delete
        End If

        Set gp_Offset = pthsRet(1)

        Set pthsRet = Nothing

End Function

Public Function gps_Array(psPathsToArray As Paths, _
                          ByVal lNumberInX As Long, ByVal dDisplacementInX As Double, _
                          ByVal lNumberInY As Long, ByVal dDisplacementInY As Double) As Paths
    
        Dim pthTempX                As Path
        Dim pthTempY                As Path
        Dim P                       As Path
        Dim pthsX                   As Paths
        Dim pthsRet                 As Paths
        Dim lngIndexX               As Long
        Dim lngIndexY               As Long
    
On Error GoTo ErrTrap
                    
        If (psPathsToArray Is Nothing) Then GoTo Controlled_Exit
                    
        ' first let's create a new path collection
        Set pthsX = Nothing
        
        With App.ActiveDrawing
                Set pthsX = .CreatePathCollection
                Set pthsRet = .CreatePathCollection
        End With
        
        Call g_LockAcam
        
        ' add the originals to collection
        For Each P In psPathsToArray
                Call pthsX.add(P)
        Next P
        
        ' look for more than one - 'cause we already have that
        If (lNumberInX > 1) Then
        
                For lngIndexX = 1 To (lNumberInX - 1)
                        
                        For Each P In psPathsToArray
                            
                                ' copy and then move the original toolpaths
                                With P
                                    
                                        Set pthTempX = .Copy
                                        Call pthTempX.MoveL((dDisplacementInX * lngIndexX), 0)
                                        
                                        ' add to collection
                                        Call pthsX.add(pthTempX)
                                    
                                End With
                            
                        Next P
                                            
                Next lngIndexX
            
        End If
        
        ' add all X to return collection
        For Each P In pthsX
                Call pthsRet.add(P)
        Next P
                        
        ' now to the Y
        If (lNumberInY > 1) Then
                
                For Each P In pthsX
                                            
                        For lngIndexY = 1 To (lNumberInY - 1)
                                                        
                                ' copy and then move the original geometry
                                With P
                                        
                                        Set pthTempY = .Copy
                                        Call pthTempY.MoveL(0, (dDisplacementInY * lngIndexY))
                                        
                                        ' add to return collection
                                        Call pthsRet.add(pthTempY)
                                        
                                End With
                            
                        Next lngIndexY
                
                Next P
                
        End If

Controlled_Exit:
        
        Set P = Nothing
        Set pthTempX = Nothing
        Set pthTempY = Nothing
        Set pthsX = Nothing
        
        Set gps_Array = pthsRet
        
        Call g_UnlockAcam

Exit Function

ErrTrap:
    
        MsgBox Err.Description, vbExclamation
        Resume Controlled_Exit

End Function

Public Function gps_Repeat(psPathsToRepeat As Paths, _
                           ByVal lNumberOfCopies As Long, _
                           ByVal dDisplacementInX As Double, _
                           ByVal dDisplacementInY As Double) As Paths
    
        Dim pthTemp                 As Path
        Dim pthsRet                 As Paths
        Dim P                       As Path
        Dim lngIndex                As Long
    
On Error GoTo ErrTrap
            
        If (psPathsToRepeat Is Nothing) Then GoTo Controlled_Exit
        
        Set pthsRet = App.ActiveDrawing.CreatePathCollection
        
        Call g_LockAcam
        
        ' add original paths to return collection
        For Each P In psPathsToRepeat
                Call pthsRet.add(P)
        Next P
        
        ' look for more than one - 'cause we already have that
        If (lNumberOfCopies > 1) Then
        
                For lngIndex = 1 To (lNumberOfCopies - 1)
  
                        ' loop thru all toolpaths and move them as well
                        For Each P In psPathsToRepeat
                        
                                With P
                                        Set pthTemp = .Copy
                                        Call pthTemp.MoveL((dDisplacementInX * lngIndex), (dDisplacementInY * lngIndex))
                                        Call pthsRet.add(pthTemp)
                                End With
                            
                        Next P
                
                Next lngIndex
        
        End If
 
Controlled_Exit:
        
        Set gps_Repeat = pthsRet
        
        Set pthsRet = Nothing
        Set pthTemp = Nothing
        Set P = Nothing
        
        Call g_UnlockAcam
        
Exit Function

ErrTrap:
    
        MsgBox Err.Description, vbExclamation
        Resume Controlled_Exit

End Function

Public Sub g_ExtendByDistance(P As Path, ByVal dDistance As Double, ByVal bIsPercent As Boolean, ByVal iEndToExtend As AlphaEndToExtend)

        Dim Drw                     As Drawing
        Dim pthRemainder            As Path
        Dim pthBoundary             As Path
        Dim E                       As Element
        Dim E2                      As Element
        Dim xyzPt                   As POINT_XYZ
        Dim xyzInc                  As POINT_XYZ
        Dim dblLen                  As Double
        Dim dblPercent              As Double
        Dim dblRd                   As Double
        Dim intPolarity             As Integer
        Dim blnReverse              As Boolean

        If (P Is Nothing) Then GoTo Controlled_Exit
        If (dDistance = 0) Then GoTo Controlled_Exit

        Set Drw = App.ActiveDrawing
        
        blnReverse = False
        dblLen = P.Length
        
        ' reverse geo?
        If (iEndToExtend = alphaExtend_START) Then
                blnReverse = True
                Call P.Reverse
        End If

        Set E = P.GetLastElem
        
        With E
    
                If .IsArc Then
                        
                        intPolarity = IIf(.CW, -1, 1)
                
                        If bIsPercent Then
                                dblPercent = (.IncludedAngle * (dDistance / 100))
                        Else
                                dblPercent = ((dDistance / (Pi * .Radius * 2)) * 360)
                        End If
                
                        If (dblPercent > 180) Then dblPercent = 180
                        
                        If (dDistance > 0) Then
                                Set pthBoundary = Drw.Create2DLine(.CenterXL, .CenterYL, .EndXL, .EndYL)
                                Call pthBoundary.RotateL((dblPercent * intPolarity), .CenterXL, .CenterYL)
                        End If
            
                Else
    
                        If bIsPercent Then
                                dblPercent = ((dDistance / 100) + 1)
                        Else
                                dblPercent = ((dDistance / .Length) + 1)
                        End If
        
                        If (dDistance > 0) Then
                                
                                xyzInc.X = (.EndXL - .StartXL)
                                xyzInc.Y = (.EndYL - .StartYL)
                                
                                Set pthBoundary = Drw.Create2DLine(.EndXL, .EndYL, _
                                                                 (.StartXL + (xyzInc.X * dblPercent)), _
                                                                 (.StartYL + (xyzInc.Y * dblPercent)))
                                                                 
                                Call pthBoundary.RotateL(90, (.StartXL + (xyzInc.X * dblPercent)), _
                                                             (.StartYL + (xyzInc.Y * dblPercent)))
                        End If
        
                End If
    
                If (dDistance < 0) Then
        
                        If bIsPercent Then
                                dblRd = (.Length * (dDistance / 100))
                        Else
                                dblRd = dDistance
                        End If
                
                        If ((P.Length - dblRd) < 0) Then GoTo Controlled_Exit
                
                        Call P.PointAtDistanceAlongPathL((P.Length - dblRd), xyzPt.X, xyzPt.Y, E2)
                        Set pthRemainder = P.BreakAtPoint(xyzPt.X, xyzPt.Y)
                
                        pthRemainder.Selected = True
        
                Else
                        pthBoundary.Selected = True
                        Call P.Extend(False)
                End If
                
        End With

        Call Drw.DeleteSelected
        
        ' reset direction?
        If blnReverse Then Call P.Reverse
    
Controlled_Exit:
        
        Set E = Nothing
        Set E2 = Nothing
        Set pthRemainder = Nothing
        Set pthBoundary = Nothing
        Set Drw = Nothing

Exit Sub
        
End Sub

Public Sub g_PrintAttributsToDebug(ByVal Obj As Object, Optional sLabel As String = vbNullString)
        
        Dim lngAtt                  As Long
        Dim lngAtts                 As Long
        Dim strAtt                  As String

On Error GoTo ErrTrap
        
        If (Obj Is Nothing) Then Exit Sub

        lngAtts = Obj.GetAttributeCount
        
        If (lngAtts = 0) Then
                Debug.Print "-> " & sLabel & " (0) <-"
                Debug.Print ""
                Exit Sub
        End If
        
        Debug.Print "-> " & sLabel & " (" & lngAtts & ")"
        Debug.Print ""
        
        For lngAtt = 1 To lngAtts
                strAtt = Obj.GetAttributeName(lngAtt)
                Debug.Print "   " & strAtt & " = " & Obj.Attribute(strAtt)
        Next lngAtt
        
        Debug.Print ""
        Debug.Print "<-"
                                        
Controlled_Exit:
                                        
Exit Sub

ErrTrap:

        Resume Controlled_Exit

End Sub

Public Sub g_CopyObjectAtts(oFrom As Object, oTo As Object, Optional ByVal bOverwrite As Boolean = True)
        
        Dim lngAtts                 As Long
        Dim lngAtt                  As Long
        Dim strAtt                  As String
        
On Error Resume Next
        
        lngAtts = oFrom.GetAttributeCount
        
        For lngAtt = 1 To lngAtts
                
                strAtt = oFrom.GetAttributeName(lngAtt)
                
                ' if already exists, look to overwrite
                If gb_AttrExists(strAtt, oTo) Then
                        If bOverwrite Then oTo.Attribute(strAtt) = oFrom.Attribute(strAtt)
                Else
                        oTo.Attribute(strAtt) = oFrom.Attribute(strAtt)
                End If
        
        Next lngAtt

Controlled_Exit:
        
Exit Sub

End Sub

Public Sub g_CopyPathAtts(pFrom As Path, pTo As Path, Optional ByVal bOverwrite As Boolean = True, Optional ByVal bIncludeElements As Boolean = True)
        
        Dim elmTo                   As Element
        Dim elmFrom                 As Element
        Dim lngE                    As Long
        Dim lngECount               As Long
        Dim lngAtts                 As Long
        Dim lngAtt                  As Long
        Dim strAtt                  As String
        
On Error Resume Next
        
        lngAtts = pFrom.GetAttributeCount
        
        For lngAtt = 1 To lngAtts
                
                strAtt = pFrom.GetAttributeName(lngAtt)
                
                ' if already exists, look to overwrite
                If gb_AttrExists(strAtt, pTo) Then
                        If bOverwrite Then pTo.Attribute(strAtt) = pFrom.Attribute(strAtt)
                Else
                        pTo.Attribute(strAtt) = pFrom.Attribute(strAtt)
                End If
        
        Next lngAtt
        
        If bIncludeElements Then
                
                lngECount = pFrom.Elements.count
                
                If (lngECount = pTo.Elements.count) Then
                
                        For lngE = 1 To lngECount
                                
                                Set elmFrom = pFrom.Elements(lngE)
                                Set elmTo = pTo.Elements(lngE)
                                
                                If Not (elmTo Is Nothing) Then
                                        Call g_CopyElementAtts(elmFrom, elmTo)
                                End If
                        
                        Next lngE
                
                End If
                
        End If

Controlled_Exit:

        Set elmTo = Nothing
        Set elmFrom = Nothing
        
Exit Sub

End Sub

Public Sub g_CopyElementAtts(eFrom As Element, eTo As Element, Optional ByVal bOverwrite As Boolean = True)

        Dim lngAtts                 As Long
        Dim lngAtt                  As Long
        Dim strAtt                  As String
        
        lngAtts = eFrom.GetAttributeCount
        
        For lngAtt = 1 To lngAtts
                
                strAtt = eFrom.GetAttributeName(lngAtt)
                
                ' if already exists, look to overwrite
                If gb_AttrExists(strAtt, eTo) Then
                        If bOverwrite Then eTo.Attribute(strAtt) = eFrom.Attribute(strAtt)
                Else
                        eTo.Attribute(strAtt) = eFrom.Attribute(strAtt)
                End If
        
        Next lngAtt

End Sub

Public Sub g_ClearAttributes(oAcamObject As Object)
                                  
        Dim lngAtt                  As Long
        Dim lngAtts                 As Long
                                                      
On Error GoTo ErrTrap
                                   
        With oAcamObject
                
                lngAtts = .GetAttributeCount
                
                For lngAtt = lngAtts To 1 Step -1
                        Call .DeleteAttribute(.GetAttributeName(lngAtt))
                Next lngAtt
        
        End With
                                                                   
Controlled_Exit:
        
        If (Err.Number <> 0) Then Call Err.Clear
    
Exit Sub

ErrTrap:
        
        Resume Controlled_Exit

End Sub

Public Sub g_AppendPathsToCollection(psSource As Paths, psToAdd As Paths)
        
        Dim P                       As Path
    
On Error Resume Next
        
        For Each P In psToAdd
                
                Call psSource.add(P)
                
                ' if path already existed in collection, an error would be raised
                If (Err.Number <> 0) Then Call Err.Clear
        
        Next P
    
        Set P = Nothing

End Sub

Public Sub g_InsertAndDrag2DObject(Optional ByVal iToInsertAndDrag As AlphaDragObjects = alphaDrag_BOTH)
        
        ' !! will only drag 2D objects into position !!
        
        Dim Drw                     As Drawing
        Dim drwTemp                 As Drawing
        Dim pthTemp                 As Path
        Dim pthsFromTemp            As Paths
        Dim pthsTemp                As Paths
        Dim strFileName             As String
        Dim strFullName             As String
        Dim dblX                    As Double
        Dim dblY                    As Double
        Dim lngRet                  As Long
        
On Error GoTo ErrTrap
        
        ' set undo point
        With App
                Call .SetUndoCommandName("Insert and Drag")
                Call .SetUndoPoint
        End With
        
        ' snag the active drawing
        Set Drw = App.ActiveDrawing
        
        ' prompt user to select drawing to insert - bail if canceled
        If Not App.GetAlphaCamFileName("Select Drawing to Insert", acamFileTypeDRAWING, acamFileActionOPEN, strFullName, strFileName) Then
                GoTo Controlled_Exit
        End If
                                                
        ' open selected drawing as temporary
        Set drwTemp = App.OpenTempDrawing(strFullName)
        
        ' create new path collecion to hold temp paths to drag
        Set pthsTemp = Drw.CreatePathCollection
        
        ' snag the geometries within the temp drawing if needed
        If (iToInsertAndDrag And alphaDrag_GEOS) Or (iToInsertAndDrag And alphaDrag_BOTH) Then
        
                Set pthsFromTemp = drwTemp.Geometries
                
                ' loop thru all geos within temp drawing and add to temp collection
                For Each pthTemp In pthsFromTemp
                        Call pthsTemp.add(pthTemp)
                Next pthTemp
                        
        End If
        
        ' now snag the toolpaths within the temp drawing if needed
        If (iToInsertAndDrag And alphaDrag_TOOLPATHS) Or (iToInsertAndDrag And alphaDrag_BOTH) Then
        
                Set pthsFromTemp = drwTemp.ToolPaths
        
                ' loop thru all toolpaths within temp drawing and add to temp collection
                For Each pthTemp In pthsFromTemp
                        Call pthsTemp.add(pthTemp)
                Next pthTemp
                        
        End If
        
        ' make sure we've got something to drag
        If (pthsTemp.count = 0) Then GoTo Controlled_Exit
        
        ' prompt user to drag into position
        lngRet = pthsTemp.DragMove("Drag to Position", 0, 0, 0, Nothing, dblX, dblY)
        
        Debug.Print lngRet; dblX; dblY
                
        ' will return 0 if drag completed, 1 if user cancels
        If (lngRet = 0) Then
                
                ' loop thru all temp paths, move to active drawing and then into location
                For Each pthTemp In pthsTemp
                        With pthTemp
                                Call .MoveToDrawing(Drw)
                                Call .MoveL(dblX, dblY)
                        End With
                Next pthTemp
        
        End If

Controlled_Exit:
        
        ' clean up
        Set Drw = Nothing
        Set drwTemp = Nothing
        Set pthsTemp = Nothing
        Set pthsFromTemp = Nothing
        Set pthTemp = Nothing

Exit Sub

ErrTrap:
        
        MsgBox Err.Description, vbExclamation
        Resume Controlled_Exit

End Sub

Public Function gb_OffsetToLayer(psPathsToOffset As Paths, psPathsOffset As Paths, oLayer As Layer, _
                                 ByVal dOffsetDistance As Double, ByVal eOffsetSide As AcamToolSide) As Boolean

        Dim Drw                     As Drawing
        Dim pthOff                  As Path
        Dim pthsOff                 As Paths
        Dim pthToOffset             As Path
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap

        ' assume not
        blnRet = False
        
        ' anything picked?
        If (psPathsToOffset Is Nothing) Then GoTo Controlled_Exit
        
        ' just make sure we've got the layer
        If (oLayer Is Nothing) Then
                MsgBox "Unable to locate layer.", vbInformation, App.name
                GoTo Controlled_Exit
        End If
        
        ' create return collection
        Set psPathsOffset = Drw.CreatePathCollection
        
        ' now offset the originals and move to the desired layer
        For Each pthToOffset In psPathsToOffset
                
                Set pthsOff = pthToOffset.Offset(dOffsetDistance, eOffsetSide)
        
                For Each pthOff In pthsOff
                        Call pthOff.SetLayer(oLayer)
                        Call psPathsOffset.add(pthOff)
                Next pthOff
                
        Next pthToOffset

        Drw.Redraw
        DoEvents

        ' looks like we made it!
        blnRet = True

Controlled_Exit:

        Set Drw = Nothing
        Set pthToOffset = Nothing
        Set pthOff = Nothing
        Set pthsOff = Nothing
        
        gb_OffsetToLayer = blnRet
        
Exit Function

ErrTrap:
        
        MsgBox Err.Description, vbExclamation
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gb_CopyToLayer(psPathsToCopy As Paths, psPathsCopied As Paths, oLayer As Layer) As Boolean

        Dim Drw                     As Drawing
        Dim pthToCopy               As Path
        Dim pthCopy                 As Path
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap

        ' assume not
        blnRet = False

        Set Drw = App.ActiveDrawing
                
        ' anything picked?
        If (psPathsToCopy Is Nothing) Then GoTo Controlled_Exit
        
        ' just make sure we've got the layer
        If (oLayer Is Nothing) Then
                MsgBox "Unable to locate layer.", vbInformation, App.name
                GoTo Controlled_Exit
        End If
                
        ' create return collection
        Set psPathsCopied = Drw.CreatePathCollection
        
        ' now copy the originals and move to the desired layer
        For Each pthToCopy In psPathsToCopy
                
                ' copy it and set the layer
                Set pthCopy = pthToCopy.Copy
                Call pthCopy.SetLayer(oLayer)
                Call psPathsCopied.add(pthCopy)
                                
        Next pthToCopy

        Drw.Redraw
        DoEvents

        ' looks like we made it!
        blnRet = True

Controlled_Exit:

        Set Drw = Nothing
        Set pthCopy = Nothing
        Set pthToCopy = Nothing
        
        gb_CopyToLayer = blnRet
    
Exit Function

ErrTrap:
        
        MsgBox Err.Description, vbExclamation
        blnRet = False
        Resume Controlled_Exit

End Function

Public Function gb_MoveToLayer(psPathsToMove As Paths, oLayer As Layer) As Boolean

        Dim Drw                     As Drawing
        Dim pthToMove               As Path
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap

        ' assume not
        blnRet = False

        Set Drw = App.ActiveDrawing
                
        ' anything picked?
        If (psPathsToMove Is Nothing) Then GoTo Controlled_Exit
        
        ' just make sure we've got the layer
        If (oLayer Is Nothing) Then
                MsgBox "Unable to locate layer.", vbInformation, App.name
                GoTo Controlled_Exit
        End If
        
        ' now move to the desired layer
        For Each pthToMove In psPathsToMove
                Call pthToMove.SetLayer(oLayer)
        Next pthToMove

        Drw.Redraw
        DoEvents

        ' looks like we made it!
        blnRet = True

Controlled_Exit:

        Set Drw = Nothing
        Set pthToMove = Nothing
        
        gb_MoveToLayer = blnRet
    
Exit Function

ErrTrap:
        
        MsgBox Err.Description, vbExclamation
        blnRet = False
        Resume Controlled_Exit

End Function

Public Sub g_SetUnsetOpenElement(E As Element, Optional ByVal iSetMethod As AlphaSetUnsetOpenElementMethod = alphaOpenE_AUTO)
        
        Const LicomUKDMBElementBoundaryType         As String = "LicomUKDMBElementBoundaryType"
        
        If (E Is Nothing) Then Exit Sub
        
        Select Case iSetMethod
                
                Case alphaOpenE_AUTO
                        
                        If (E.Attribute(LicomUKDMBElementBoundaryType) <> 0) Then
                                Call E.DeleteAttribute(LicomUKDMBElementBoundaryType)
                        Else
                                E.Attribute(LicomUKDMBElementBoundaryType) = 1
                        End If
        
                Case alphaOpenE_SET: E.Attribute(LicomUKDMBElementBoundaryType) = 1
        
                Case alphaOpenE_UNSET: Call E.DeleteAttribute(LicomUKDMBElementBoundaryType)
        
        End Select
        
        Call E.Redraw(acamDrawTypeNORMAL)

End Sub

Public Sub g_DisableGeos(Optional psGeos As Paths = Nothing, Optional ByVal bSelectedOnly As Boolean = False)
        
        Dim P                       As Path
        Dim PS                      As Paths
        
        If (psGeos Is Nothing) Then
                Set PS = App.ActiveDrawing.Geometries
        Else
                Set PS = psGeos
        End If
        
        For Each P In PS
                
                If bSelectedOnly Then
                        If P.Selected Then
                                P.Disabled = True
                                Call P.Redraw
                        End If
                Else
                        P.Disabled = True
                        Call P.Redraw
                End If
                
        Next P
        
        Set P = Nothing
        Set PS = Nothing

End Sub

Public Sub g_DisableAllGeosExcept(psKeepEnabled As Paths)
        
        Dim P1                      As Path
        Dim P2                      As Path
        Dim blnDisable              As Boolean
        
        If (psKeepEnabled Is Nothing) Then Exit Sub
        
        For Each P1 In App.ActiveDrawing.Geometries
                
                blnDisable = True
                
                For Each P2 In psKeepEnabled
                                                     
                        If P1.IsSame(P2) Then
                                blnDisable = False
                                Exit For
                        End If
                
                Next P2
                
                If blnDisable Then
                        P1.Disabled = True
                        Call P1.Redraw
                End If
        
        Next P1

        Set P1 = Nothing
        Set P2 = Nothing

End Sub

Public Sub g_EnableGeos(Optional psGeos As Paths = Nothing, Optional ByVal bSelectedOnly As Boolean = False)
        
        Dim P                       As Path
        Dim PS                      As Paths
        
        If (psGeos Is Nothing) Then
                Set PS = App.ActiveDrawing.Geometries
        Else
                Set PS = psGeos
        End If
        
        For Each P In PS
                
                If bSelectedOnly Then
                        If P.Selected Then
                                P.Disabled = False
                                Call P.Redraw
                        End If
                Else
                        P.Disabled = False
                        Call P.Redraw
                End If
                
        Next P
        
        Set P = Nothing
        Set PS = Nothing

End Sub

Public Sub g_EnableAllGeosExcept(psKeepDisabled As Paths)
        
        Dim P1                      As Path
        Dim P2                      As Path
        Dim blnEnable               As Boolean
        
        If (psKeepDisabled Is Nothing) Then Exit Sub
        
        For Each P1 In App.ActiveDrawing.Geometries
                
                blnEnable = True
                
                For Each P2 In psKeepDisabled
                                                     
                        If P1.IsSame(P2) Then
                                blnEnable = True
                                Exit For
                        End If
                
                Next P2
                
                If blnEnable Then
                        P1.Disabled = False
                        Call P1.Redraw
                End If
        
        Next P1

        Set P1 = Nothing
        Set P2 = Nothing

End Sub

Public Sub g_Redraw(Optional ByVal bZoomAll As Boolean = False, Optional bRefresh As Boolean = True)

        Dim VW                      As ViewWindow

        Const DEF_REFRESH           As Long = 33619

On Error Resume Next
        
        ' force redraw and refresh if API version is OK
        With App
        
                If bZoomAll Then Call .ActiveDrawing.ZoomAll
                                
                ' redraw all view windows
                For Each VW In App.ActiveDrawing.ViewWindows
                        Call VW.Redraw
                Next VW
                                
                If (.ApiVersion >= 20040928) Then
                        If bRefresh Then Call .Frame.RunCommand(DEF_REFRESH)
                End If
        
        End With
        
        Set VW = Nothing
        
End Sub

Public Sub g_LockAcam()
        
        Dim VW                      As ViewWindow
        Dim lngRet                  As Long
        
On Error Resume Next

        With App
                
                .ActiveDrawing.ScreenUpdating = False
                
                With .Frame
                
                        .ProjectBarUpdating = False
                                                        
                        ' prevent redraw for all view windows
                        'For Each VW In App.ActiveDrawing.ViewWindows
                        '        lngRet = SendMessage(VW.ViewWindowHandle, WM_SETREDRAW, False, 0&)
                        'Next VW
                                                
                End With
                
        End With
        
        Set VW = Nothing
        
End Sub

Public Sub g_UnlockAcam(Optional ByVal bZoomAll As Boolean = False)
        
        Dim VW                      As ViewWindow
        Dim lngRet                  As Long
        
On Error Resume Next
            
        ' prevent redraw for all view windows
        'For Each VW In App.ActiveDrawing.ViewWindows
        '        lngRet = SendMessage(VW.ViewWindowHandle, WM_SETREDRAW, True, 0&)
        'Next VW
                        
        With App
                .Frame.ProjectBarUpdating = True
                .ActiveDrawing.ScreenUpdating = True
        End With
        
        Call g_Redraw(bZoomAll)
        
End Sub

Public Function gs_ReadAcamCTX(ByVal lDollar As Long, ByVal lIndex As Long, ByVal sDefault As String)

        Dim strRet                  As String
        Dim strCTX                  As String
        
        strCTX = ms_AcamCTX
        
        strRet = gv_CTX(lDollar, lIndex, sDefault, strCTX)
        
        gs_ReadAcamCTX = strRet

End Function

Public Function gs_ReadAeditCTX(ByVal lDollar As Long, ByVal lIndex As Long, ByVal sDefault As String)

        Dim strRet                  As String
        Dim strCTX                  As String
        
        strCTX = ms_AeditCTX
        
        strRet = gv_CTX(lDollar, lIndex, sDefault, strCTX)
        
        gs_ReadAeditCTX = strRet

End Function

Public Function gs_ReadAcamNestCTX(ByVal lDollar As Long, ByVal lIndex As Long, ByVal sDefault As String)

        Dim strRet                  As String
        Dim strCTX                  As String
        
        strCTX = ms_AcamNestCTX
        
        strRet = gv_CTX(lDollar, lIndex, sDefault, strCTX)
        
        gs_ReadAcamNestCTX = strRet

End Function

Public Function gv_CTX(ByVal lDollar As Long, ByVal lIndex As Long, vDefault As Variant, _
                     Optional ByVal sFileToRead As String = vbNullString, _
                     Optional ByVal eVarType As AlphaVariableType = alphaVarType_STRING) As Variant
        
        Const DEF_BYPASS_ERR        As Long = 2
        
        Dim strFile                 As String
        Dim strBuf                  As String
        Dim strRet                  As String
                
        strRet = CStr(vDefault)
        
        ' build the path to the text file
        If (sFileToRead <> vbNullString) Then
                strFile = sFileToRead
        Else
                strFile = DEF_CTX
        End If
        
        strBuf = App.Frame.ReadTextFile2(strFile, lDollar, lIndex, DEF_BYPASS_ERR)
                        
        ' ReadTextFile2 will return a null string if not found
        If (Len(Trim$(strBuf)) > 0) Then
                strRet = strBuf
        Else
                ' look for language tag at end of string - is added sometimes at design time
                Select Case True
                        
                        Case (Len(strRet) < 3)      ' do nothing
                        Case (Right$(Trim$(strRet), 3) = "!!!"), _
                             (Right$(Trim$(strRet), 3) = "@@@")
                        
                                strRet = Left$(strRet, (Len(Trim$(strRet)) - 3))
                                
                End Select
                
        End If
                
        If (Err.Number <> 0) Then Err.Clear
        
        ' set return value
        Select Case eVarType
                Case alphaVarType_BOOLEAN: gv_CTX = CBool(strRet)
                Case alphaVarType_SINGLE: gv_CTX = CSng(strRet)
                Case alphaVarType_DOUBLE: gv_CTX = CDbl(strRet)
                Case alphaVarType_INTEGER: gv_CTX = CInt(strRet)
                Case alphaVarType_LONG: gv_CTX = CLng(strRet)
                Case alphaVarType_STRING: gv_CTX = CStr(strRet)
                Case Else: gv_CTX = strRet
        End Select
        
End Function

Public Function gl_AcamHwnd() As Long
        gl_AcamHwnd = App.Frame.WindowHandle    '' FindWindow32(vbNullString, App.Name)
End Function

Public Function gl_GetAvailablePostAttNumber() As Long
        
        Dim lngIndex                As Long
        Dim lngCount                As Long
        Dim lngAtt                  As Long
        Dim lngRet                  As Long
        
On Error Resume Next
                
        lngRet = 1
        
        lngCount = App.GetPostAttributeCount
        
        For lngIndex = 1 To lngCount
                
                lngAtt = App.GetPostAttributeNumber(lngIndex)
                
                If (lngAtt > lngRet) Then lngRet = (lngAtt + 1)
                If (Err.Number <> 0) Then Err.Clear
                
        Next lngIndex

        gl_GetAvailablePostAttNumber = lngRet
        
End Function

Public Function gl_NestExtensionIndex(nExtensions As Object, ByVal sName As String) As Long
        
        ' nExtensions   = App.Nesting.NestExtensions object containing Extension to be checked
        ' sName         = Name of Extension to be checked
        
        Dim NE                      As Object   ' ACAMNESTLib.NestExtension
        Dim lngID                   As Long
        Dim lngRet                  As Long
                
        lngRet = 0
        lngID = 1
        
        If Not (nExtensions Is Nothing) Then
        
                For Each NE In nExtensions
                        
                        If (StrComp(NE.name, sName, vbTextCompare) = 0) Then
                                lngRet = lngID
                                Exit For
                        End If
                        
                        lngID = (lngID + 1)
                
                Next NE
        
        End If
        
        gl_NestExtensionIndex = lngRet
        
        Set NE = Nothing
        
End Function

Public Function gb_HasSTART(ByVal sFileToCheck As String) As Boolean

        Dim intFile                 As Integer
        Dim strLine                 As String
        Dim blnRet                  As Boolean
        
        Const START                 As String = "START"

On Error GoTo ErrTrap
    
        blnRet = False
    
        ' get a free file handle
        intFile = FreeFile
        
        ' overwrite the nc file
        Open sFileToCheck For Input As #intFile
            
        ' scan each line
        Do While Not EOF(intFile)
            
                ' get a line
                Line Input #intFile, strLine
               
                ' make sure we have something to test
                If Len(strLine) > 0 Then
                    
                        If (UCase$(Mid$(strLine, 1, 5)) = START) Then
                                blnRet = True
                                Exit Do
                        End If
                                    
                End If
                            
        Loop

Controlled_Exit:
        
        gb_HasSTART = blnRet
        
        Close
        
Exit Function

ErrTrap:

        blnRet = False
        MsgBox Err.Description, vbExclamation
        Resume Controlled_Exit

End Function

Private Function ms_ProgramLetter(Optional ByVal bUcase As Boolean = False, Optional sName As String = vbNullString) As String
        
        Dim strRet                  As String
        
        Select Case App.ProgramLetter
                Case Asc("M"): strRet = "m": sName = gs_ReadAeditCTX(1940, 1, "Mill")       ' mill
                Case Asc("R"): strRet = "r": sName = gs_ReadAeditCTX(1940, 2, "Router")     ' router
                Case Asc("T"): strRet = "t": sName = gs_ReadAeditCTX(1940, 3, "Turning")    ' lathe
                Case Asc("L"): strRet = "l": sName = gs_ReadAeditCTX(1940, 4, "Laser")      ' laser
                Case Asc("E"): strRet = "e": sName = gs_ReadAeditCTX(1940, 5, "Wire")       ' wire edm
                Case Asc("P"): strRet = "p": sName = gs_ReadAeditCTX(1940, 6, "Punch")      ' punch
                Case Asc("F"): strRet = "f": sName = gs_ReadAeditCTX(1940, 7, "Flame")      ' flame
                Case Asc("J"): strRet = "j": sName = gs_ReadAeditCTX(1940, 9, "Water Jet")  ' water jet
                Case Asc("S"): strRet = "s": sName = gs_ReadAeditCTX(1940, 10, "Stone")     ' marble
        End Select
        
        If bUcase Then strRet = UCase$(strRet)
        
        ms_ProgramLetter = strRet
        
End Function

Private Function ms_AcamCTX() As String
        
        Dim strRet                  As String
        
        strRet = gs_AppDir
        strRet = strRet & "acam.ctx"
        
        ms_AcamCTX = strRet
        
End Function

Private Function ms_AeditCTX() As String
        
        Dim strRet                  As String
        
        strRet = gs_AppDir
        strRet = strRet & "aedit.ctx"
        
        ms_AeditCTX = strRet
        
End Function

Private Function ms_AcamNestCTX() As String
        
        Dim strRet                  As String
        
        strRet = gs_AppDir & "Add-Ins\Nesting\"
        strRet = strRet & "acamnest.ctx"
        
        ms_AcamNestCTX = strRet
        
End Function

Public Function PSDbl(ByVal S As String) As Double

        ' This function should always be used to convert a string
        ' e.g. from a text box to a Double or Single value
        '
        ' Convert string to floating point value.
        '
        ' Uses Val at the moment, but may use CDbl in future to allow
        ' "," as decimal separator.
        '
        ' Val always use ".", CDbl uses "," or ".", depending on the Regional
        ' Settings in Control Panel. But the Alphacam Evaluate function uses
        ' "." for decimal, and "," for parameter separators in some functions.
        ' To allow "," as decimal separator, extensive changes would be needed
        ' in Alphacam, so VBA should only use "." to be consistent.
        '
        ' Also, if "." is passed to CDbl when regional settings use "," for decimal,
        ' the return value will be incorrect - returns number less any decimal.
        ' Same type of problem if a "," is passed to Val - returns only the whole number.

        S = gs_NoComma(S)
        PSDbl = Val(S)

End Function

Public Function PSTol(ByVal vVal As Variant, Optional ByVal Places As Long = 4) As Double
        
        Dim lngPlaces               As Long
        Dim strFormat               As String
        Dim strRet                  As String
        
On Error Resume Next
        
        lngPlaces = IIf((Places < 0), 0, (Places - 1))
        strFormat = "#0.0" & String$((lngPlaces), "#")
        
        strRet = Format$(PSDbl(vVal), strFormat)
        strRet = PSDbl(strRet)
        
        PSTol = Val(strRet)
                
End Function

Public Function PSStr(ByVal vVal As Variant, Optional ByVal Places As Long = 4) As String

        Dim lngPlaces               As Long
        Dim strFormat               As String
        Dim strRet                  As String
        
On Error Resume Next
        
        lngPlaces = IIf((Places < 0), 0, (Places - 1))
        strFormat = "#0.0" & String$((lngPlaces), "#")
        
        strRet = Format$(PSDbl(vVal), strFormat)
        strRet = gs_NoZeros(strRet)
        
        PSStr = strRet

End Function

-------------------------------------------------------------------------------
