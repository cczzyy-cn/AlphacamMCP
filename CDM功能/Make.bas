Option Explicit
'

Private Sub m_DeleteNestSheet(NestSheetToDelete As NestSheet)
  Dim Npi           As NestPartInstance
  Dim Tp            As Path
  Dim BobblePath    As Path
  Dim strSheetName  As String
  Dim DrawingText   As Text

  ' Delete the last sheet and components
  For Each Npi In NestSheetToDelete.Parts
    For Each Tp In Npi.Paths
      Tp.Selected = True
    Next
  Next
  NestSheetToDelete.Path.Selected = True
  
  strSheetName = GetNestedSheetNumber(NestSheetToDelete.Name, True, False)
  
  For Each DrawingText In ActiveDrawing.Text
    If DrawingText.Attribute(DEF_ATT_SHEET_ITEM_NOS) = "1" And DrawingText.Attribute(DEF_ATT_SHEET_IDENT) = strSheetName Then
      DrawingText.Selected = True
    End If
  Next
  For Each BobblePath In ActiveDrawing.Geometries
    If BobblePath.Attribute(DEF_ATT_SHEET_BOBBLE) And BobblePath.Attribute(DEF_ATT_SHEET_IDENT) = strSheetName Then
      BobblePath.Selected = True
    End If
  Next
  
  ActiveDrawing.DeleteSelected

End Sub

Private Function mbln_RunCustomMacro(MacroFilename As String, Door As CDoor) As Boolean

  Dim i                       As Integer
  Dim iL                      As Integer
  Dim iU                      As Integer
  Dim strDims()               As String
  Dim strDimsDesc()           As String
  Dim lngGeoNumber            As Long
  Dim bVBAProjectOpen         As Boolean

On Error GoTo mbln_RunCustomMacro_Error
  
  mbln_RunCustomMacro = True
  
  ' Determine the custom macro project name if required
  If gstr_CustomMacroProjectName = "" Or UCase(MacroFilename) <> UCase(gstr_CustomMacroFileName) Then
    App.LoadAddIn MacroFilename
    gstr_CustomMacroProjectName = mstr_GetCustomMacroName(MacroFilename)
  End If

  With Door
    
    .CustomMacroSuccess = True

    '..split the dimension string
    strDims = Split(.UserVariableString, ";")
    
    iL = LBound(strDims)
    iU = UBound(strDims)

    For i = iL To iU
        .UserVariablesCollection.add PDbl(Trim$(strDims(i)))
    Next i

    DoEvents
            
    '..split the dimension description string
    strDimsDesc = Split(.UserVariableDescriptionString, ";")
    
    iL = LBound(strDimsDesc)
    iU = UBound(strDimsDesc)
    
    For i = iL To iU
        .UserVariablesDescriptionsCollection.add Trim$(strDimsDesc(i))
    Next i
    
    DoEvents
            
    ' Test to see if the project is already open within VBA
    ' Loading an already open project will reset any breakpoint the user may have set
    bVBAProjectOpen = gbln_IsVBAProjectOpen(gstr_CustomMacroProjectName)
    
    If Not bVBAProjectOpen Then
      ' Load the VBA Macro
      App.LoadAddIn gstr_CustomMacroFileName
    End If
            
    ' Set the JobName property so it is available within the custom macro if needed
    Door.JobName = gstr_JobName
            
    '..launch the macro
    Call App.Run(gstr_CustomMacroProjectName & "." & DEF_CUSTOM_MACRO_MODULE_NAME & "." & DEF_CUSTOM_MACRO_PROCEDURE_NAME, Door, _
                 .UserArg_0, .UserArg_1, .UserArg_2, .UserArg_3, .UserArg_4, .UserArg_5, .UserArg_6)

    If Not bVBAProjectOpen Then
      ' Finished with the VBA Project, so disable it (unload it)
      App.EnableAddIn gstr_CustomMacroFileName, False
    End If
  
  End With
    
  If Not Door.CustomMacroSuccess Then mbln_RunCustomMacro = False: GoTo Controlled_Exit

Controlled_Exit:

    '..rinse
    Erase strDims
    
Exit Function
    
mbln_RunCustomMacro_Error:
          
    If Err.Number = -2147467259 Then
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 100) & Space(3), vbExclamation + vbMsgBoxHelpButton, _
                DEF_PROJECT_NAME, App.Frame.PathOfThisAddin & DEF_BACKSLASH & DEF_HELP_NAME_CHM, DEF_HLP_ID_Defining_a_User_Defined_Door_Style
    Else
        If (Err.Number <> 0) Then WriteError Err, True, "mbln_RunCustomMacro"
    End If
    
    mbln_RunCustomMacro = False
    Resume Controlled_Exit

End Function

Private Function mbln_DrawHandleHole(Door As CDoor, HoleSize As Double, HoleXPos As Double, HoleYPos As Double) As Boolean
  Dim pthHole As Path

  If HoleXPos >= 0 And HoleXPos <= Door.Width And HoleYPos >= 0 And HoleYPos <= Door.Length Then
    Set pthHole = ActiveDrawing.CreateCircle(HoleSize, HoleXPos, HoleYPos)
    pthHole.Selected = True
    pthHole.Attribute(DEF_ATT_HANDLE_GEO) = "1"
    mbln_DrawHandleHole = True
  End If

End Function

Private Function mbln_DrawHandleHoles(Door As CDoor, bPreview As Boolean) As Boolean
    
    Dim rstHandle               As New ADODB.Recordset
    Dim FSO                     As New Scripting.FileSystemObject
    Dim pthHole                 As Path
    Dim mstDrilling             As MillStyle
    Dim dblDatumLocation        As Double
    Dim dblAdditionalOffsetX    As Double
    Dim dblAdditionalOffsetY    As Double
    Dim dblDistBetweenHoles     As Double
    Dim dblHoleDiameter         As Double
    Dim dblDoorHeight           As Double
    Dim dblDoorWidth            As Double
    Dim dblStartPos             As Double
    Dim lngHandleOrientation    As Double
    Dim lngNumHoles             As Long
    Dim lngCount                As Long
    Dim strMachiningStyle       As String
    Dim strHandleConfigName     As String
    Dim blnHolePositionErr      As Boolean
    Dim lngTpCountPre           As Long
    Dim lngTpCountPost          As Long
    Dim lngTpCount              As Long
        
    
    
    ' No handle configuration set
    If Door.HandleID = 0 Then
      mbln_DrawHandleHoles = True
      Exit Function
    End If
        
    Set rstHandle = grst_GetHandleConfiguration(Door.HandleID)
    
    If rstHandle Is Nothing Then
      mbln_DrawHandleHoles = False
      GoTo Controlled_Exit
    End If
    
    dblDoorHeight = Door.Length
    dblDoorWidth = Door.Width
            
    dblDatumLocation = gvar_CheckNull(rstHandle.Fields!DatumLocation)
    dblAdditionalOffsetX = gvar_CheckNull(rstHandle.Fields!DatumLocationAdditionalOffsetX)
    dblAdditionalOffsetY = gvar_CheckNull(rstHandle.Fields!DatumLocationAdditionalOffsetY)
    lngHandleOrientation = gvar_CheckNull(rstHandle.Fields!HoleOrientation)
    lngNumHoles = gvar_CheckNull(rstHandle.Fields!NumberOfHoles)
    dblDistBetweenHoles = gvar_CheckNull(rstHandle.Fields!HoleSpacing)
    dblHoleDiameter = gvar_CheckNull(rstHandle.Fields!HoleSize)
    strMachiningStyle = gvar_CheckNull(rstHandle.Fields!MachiningStyle)

    Dim dblHoleDatumX As Double
    Dim dblHoleDatumY As Double
    
    Select Case dblDatumLocation
      
      Case AdoorHandleDatumLocation_BottomLeft
        dblHoleDatumX = dblAdditionalOffsetX
        dblHoleDatumY = dblAdditionalOffsetY
      
      Case AdoorHandleDatumLocation_BottomCentre
        dblHoleDatumX = dblDoorWidth / 2 + dblAdditionalOffsetX
        dblHoleDatumY = dblAdditionalOffsetY
      
      Case AdoorHandleDatumLocation_BottomRight
        dblHoleDatumX = dblDoorWidth - dblAdditionalOffsetX
        dblHoleDatumY = dblAdditionalOffsetY
      
      Case AdoorHandleDatumLocation_CentreLeft
        dblHoleDatumX = dblAdditionalOffsetX
        dblHoleDatumY = dblDoorHeight / 2 + dblAdditionalOffsetY
      
      Case AdoorHandleDatumLocation_CentreCentre
        dblHoleDatumX = dblDoorWidth / 2 + dblAdditionalOffsetX
        dblHoleDatumY = dblDoorHeight / 2 + dblAdditionalOffsetY
      
      Case AdoorHandleDatumLocation_CentreRight
        dblHoleDatumX = dblDoorWidth - dblAdditionalOffsetX
        dblHoleDatumY = dblDoorHeight / 2 + dblAdditionalOffsetY
      
      Case AdoorHandleDatumLocation_UpperLeft
        dblHoleDatumX = dblAdditionalOffsetX
        dblHoleDatumY = dblDoorHeight - dblAdditionalOffsetY
      
      Case AdoorHandleDatumLocation_UpperCentre
        dblHoleDatumX = dblDoorWidth / 2 + dblAdditionalOffsetX
        dblHoleDatumY = dblDoorHeight - dblAdditionalOffsetY
      
      Case AdoorHandleDatumLocation_UpperRight
        dblHoleDatumX = dblDoorWidth - dblAdditionalOffsetX
        dblHoleDatumY = dblDoorHeight - dblAdditionalOffsetY
    
    End Select

    If lngNumHoles = 1 Then
      
      If Not mbln_DrawHandleHole(Door, dblHoleDiameter, dblHoleDatumX, dblHoleDatumY) Then
        blnHolePositionErr = True
      End If
    Else
    
      If lngHandleOrientation = AdoorHandleOrientation_Vertical Then
            
        If (lngNumHoles / 2) = (lngNumHoles \ 2) Then
          ' Even number - holes spaced around datum
          dblStartPos = dblHoleDatumY + (((lngNumHoles - 1) * dblDistBetweenHoles) / 2)
        Else
          ' Odd number - hole positioned at datum location
          dblStartPos = dblHoleDatumY + (((lngNumHoles - 1) * dblDistBetweenHoles) / 2)
        End If
                
        For lngCount = 1 To lngNumHoles
          If Not mbln_DrawHandleHole(Door, dblHoleDiameter, dblHoleDatumX, dblStartPos) Then
            blnHolePositionErr = True
          End If
          dblStartPos = dblStartPos - dblDistBetweenHoles
        Next
      
      ElseIf lngHandleOrientation = AdoorHandleOrientation_Horizontal Then
        
        If (lngNumHoles / 2) = (lngNumHoles \ 2) Then
          ' Even number - holes spaced around datum
          dblStartPos = dblHoleDatumX + (((lngNumHoles - 1) * dblDistBetweenHoles) / 2)
        Else
          ' Odd number - hole positioned at datum location
          dblStartPos = dblHoleDatumX + (((lngNumHoles - 1) * dblDistBetweenHoles) / 2)
        End If
        
        For lngCount = 1 To lngNumHoles
          If Not mbln_DrawHandleHole(Door, dblHoleDiameter, dblStartPos, dblHoleDatumY) Then
            blnHolePositionErr = True
          End If
          dblStartPos = dblStartPos - dblDistBetweenHoles
        Next
      
      End If
    End If

    If FSO.FileExists(LicomdirPath & strMachiningStyle) Then
      Set mstDrilling = MillMachiningStyles(LicomdirPath & strMachiningStyle)
      mstDrilling.ShowProgressBox = False
      lngTpCountPre = ActiveDrawing.GetToolPathCount
      mstDrilling.ApplyToSelectedGeometries
      lngTpCountPost = ActiveDrawing.GetToolPathCount
      For lngTpCount = lngTpCountPre + 1 To lngTpCountPost
        ActiveDrawing.ToolPaths(lngTpCount).Attribute(DEF_ATT_HANDLE_TP) = "1"
      Next
    Else
      MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 143) & Space(3), vbExclamation, DEF_PROJECT_NAME
      mbln_DrawHandleHoles = False
      GoTo Controlled_Exit
    End If
    
    If blnHolePositionErr Then
      If bPreview Then
      
          MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 144) & Chr(13) & _
            Frame.ReadTextFile(clsOptions.CTXFile, 500, 145), vbInformation, DEF_PROJECT_NAME
      
      Else
    
          If MsgBox(Door.TypeName & Chr(13) & Chr(13) & _
            Frame.ReadTextFile(clsOptions.CTXFile, 500, 144) & Chr(13) & _
            Frame.ReadTextFile(clsOptions.CTXFile, 500, 145) & Chr(13) & Chr(13) & _
            Frame.ReadTextFile(clsOptions.CTXFile, 500, 20), vbQuestion + vbYesNo, DEF_PROJECT_NAME) = vbNo Then
            
            mbln_DrawHandleHoles = False
            GoTo Controlled_Exit
          
          End If
      
      End If
      
    End If
    
    mbln_DrawHandleHoles = True
    
Controlled_Exit:

    On Error Resume Next
    
    '..clean up
    If rstHandle.State = adStateOpen Then
      rstHandle.Close
    End If
    Set rstHandle = Nothing
    
    Set pthHole = Nothing

Exit Function

mbln_DrawHandleHoles_Error:

  MsgBox Err.Description, vbExclamation, "mbln_DrawHandleHoles"
  If (Err.Number <> 0) Then WriteError Err, True, "mbln_DrawHandleHoles"
  mbln_DrawHandleHoles = False
  Resume Controlled_Exit

End Function

Private Sub m_DrawNestZones(Material As CMaterial, ZonesCollection As Collection)

  Dim NestZone  As CNestingZone
  Dim Door      As CDoor
  Dim pthZone   As Path
'
  For Each Door In Material.colDoors
    If Door.NestingZone <> 0 Then
      If gbln_ItemExistsInCollection(ZonesCollection, "k" & Door.NestingZoneID) Then
        Set NestZone = ZonesCollection("k" & Door.NestingZoneID)
        With NestZone
          ' TFS#65726 - Only draw the auto zone if it is used
          If (.ZoneAssignMethod = adoorNestZone_Automatic And .AutoZoneInUse) Or .ZoneAssignMethod = adoorNestZone_Manual Then
            ' TFS#65726 - draw the zone once only
            If Not .ZoneDrawn Then
              Set pthZone = ActiveDrawing.CreateRectangle(.ZoneStartX, .ZoneStartY, .ZoneStartX + .ZoneWidth, .ZoneStartY + .ZoneHeight)
              pthZone.Attribute(DEF_ATT_NEST_ZONE) = CLng(.ZoneNumber)
              pthsNestZones.add pthZone
              .ZoneDrawn = True
            End If
          End If
        End With
      End If
    End If
  Next

End Sub




Public Sub m_NCHopsOutput(SheetLength As String, _
                           SheetWidth As String, _
                           SheetThickness As String, _
                           SheetMaterial As CMaterial)

    Dim Length As String
    Dim Width As String
    Dim Thickness As String
    Dim XShift As String
    Dim YShift As String
    Dim ZShift As String
    Dim Comment As String
    Dim b5Axis As String
    Dim bAcamToolValues As String
    Dim Language As String
    Dim colSheetFilenames As New Collection
    Dim strCurrentDrawing   As String
        
    ' Set from nested sheet
    ' Swap Length and widths over
    Length = SheetWidth
    Width = SheetLength
    
    Thickness = SheetThickness
        
    ' From CDM options
    XShift = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "NCHops_OffsetX", "0")
    YShift = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "NCHops_OffsetY", "0")
    ZShift = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "NCHops_OffsetZ", "0")
    
    If ZShift = "0" Then
      ZShift = "-" & Thickness
    End If
    
    bAcamToolValues = Abs(GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "NCHOps_ToolValuesFromAcam", "0"))
    
    ' Default
    Comment = ""
    b5Axis = "0"
    Language = "NcHopsPostCommandsEng"

    Dim MyVBA As Object 'VBE '
    Dim Project As Object 'VBProject '
    Dim strProject As String
    
    Set MyVBA = App.VBE
    
    For Each Project In MyVBA.VBProjects
        If InStr(1, Project.Name, "NcHopsPost", vbTextCompare) = 1 Then
            strProject = "NcHopsPost.Make.CallFromExternal"
            Exit For
        End If
    Next

    If strProject = "" Then
    
      MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 600, 212) & Space(3), vbExclamation, DEF_PROJECT_NAME
      Exit Sub
    
    End If
    
    strCurrentDrawing = ActiveDrawing.FullName
        
    g_SaveSheets SheetMaterial.MaterialName, colSheetFilenames
    
    Dim lngCount            As Long
    Dim strDrawingName      As String
    Dim strNCName           As String
    Dim Ni                  As NestInformation
    Dim SheetPath           As Path
    Dim SheetGeo            As Path
    Dim SheetTP             As Path
    Dim SheetText           As Text
    Dim dblSheetLowerLeftX  As Double
    Dim dblSheetLowerLeftY  As Double
    
    For lngCount = 1 To colSheetFilenames.Count
    
          strDrawingName = colSheetFilenames(lngCount)
          App.OpenDrawing strDrawingName
      
          ' Only move nested sheets 2 and onwards...
          If lngCount <> 1 Then
          
              ' Move the sheet to (0,0)
              Set Ni = ActiveDrawing.GetNestInformation
              
              Set SheetPath = Ni.Sheets(1).Path
              
              dblSheetLowerLeftX = SheetPath.MinXL
              dblSheetLowerLeftY = SheetPath.MinYL
          
              For Each SheetGeo In ActiveDrawing.Geometries
                SheetGeo.MoveL -dblSheetLowerLeftX, -dblSheetLowerLeftY
              Next
              
              For Each SheetTP In ActiveDrawing.ToolPaths
                SheetTP.MoveL -dblSheetLowerLeftX, -dblSheetLowerLeftY
              Next
              
              For Each SheetText In ActiveDrawing.Text
                SheetText.MoveL -dblSheetLowerLeftX, -dblSheetLowerLeftY
              Next
          
              ActiveDrawing.AutoScale   ' .ZoomAll
              ActiveDrawing.Save
              
          End If
          
          strNCName = gstr_StripExtension(strDrawingName) & DEF_EXTENSION_HOP
          
          App.Run strProject, _
                              Length & ";" & _
                              Width & ";" & _
                              Thickness & ";" & _
                              XShift & ";" & _
                              YShift & ";" & _
                              ZShift & ";" & _
                              Comment & ";" & _
                              b5Axis & ";" & _
                              bAcamToolValues & ";" & _
                              Language & ";" & _
                              strNCName

    Next

    App.OpenDrawing strCurrentDrawing


End Sub


Public Sub m_NCHopsOutputTemp()

    Dim Length As String
    Dim Width As String
    Dim Thickness As String
    Dim XShift As String
    Dim YShift As String
    Dim ZShift As String
    Dim Comment As String
    Dim b5Axis As String
    Dim bAcamToolValues As String
    Dim Language As String
    Dim HopsFileName As String
        
    ' Set from nested sheet
    Length = "1200"
    Width = "2000"
    Thickness = "16"
        
    XShift = "0"
    YShift = "0"
    ZShift = "0"
    bAcamToolValues = "0"
    
    ' Default
    Comment = ""
    b5Axis = "0"
    Language = "NcHopsPostCommandsEng"

    Dim MyVBA As Object 'VBE '
    Dim Project As Object 'VBProject '
    Dim strProject As String
    
    
    
    Set MyVBA = App.VBE
    
    For Each Project In MyVBA.VBProjects
        If InStr(1, Project.Name, "NcHopsPost", vbTextCompare) = 1 Then
            strProject = "NcHopsPost.Make.CallFromExternal"
            Exit For
        End If
    Next

    HopsFileName = "C:\hopstest.hop"
          
    App.Run strProject, _
                        Length & ";" & _
                        Width & ";" & _
                        Thickness & ";" & _
                        XShift & ";" & _
                        YShift & ";" & _
                        ZShift & ";" & _
                        Comment & ";" & _
                        b5Axis & ";" & _
                        bAcamToolValues & ";" & _
                        Language & ";" & _
                        HopsFileName


End Sub

Private Sub m_PackFinalNestSheetToLHS(NestMaterial As CMaterial)

  Dim N                   As Nesting
  Dim Ni                  As NestInformation
  Dim Nl                  As Nestlist
  Dim NestErrorsNl        As Nestlist
  Dim Sl                  As Sheetlist
  Dim Nsh                 As NestSheet
  Dim Npi                 As NestPartInstance
  Dim NstSheetPath        As Path
  Dim Tp                  As Path
  Dim pthConsGeo          As Path
  Dim lngSheetCount       As Long
  Dim strNestFile         As String
  Dim dblRotationAngle    As Double
  Dim strNestErrors       As String
  Dim dblDrawingExtentY   As Double
  Dim cNestRect           As CNest
  Dim cNestTrue           As CNest
  Dim dblLeftmostTPRect   As Double
  Dim dblLeftmostTPTrue   As Double
  Dim dblLeftMost         As Double
  Dim dblMaxLeftmost      As Double
  Dim intNumSheetsPre     As Integer
  Dim intNumSheetsPost    As Integer
  Dim intSheetCount       As Integer
  Dim strLastSheetName    As String

  On Error GoTo m_PackFinalNestSheetToLHS_Error
  
  Set Ni = ActiveDrawing.GetNestInformation
  
  lngSheetCount = Ni.Sheets.Count
  
  Set Nsh = Ni.Sheets(lngSheetCount)
  strLastSheetName = Nsh.Name
  
  NestMaterial.PackTo = 0
  
  Set cNestRect = New CNest
  Set cNestTrue = New CNest
  
  cNestTrue.StartNestListTrueRouter gstr_JobName, NestMaterial, False, True
  cNestRect.StartNestListRectRouter gstr_JobName, NestMaterial, True
  
  For Each Npi In Nsh.Parts
    For Each Tp In Npi.Paths
      If Tp.IsToolPath Then
        strNestFile = Tp.Attribute(DEF_ATT_ARD)
        'dblRotationAngle = CDbl(Tp.Attribute(DEF_ATT_ROT_ANGLE))
        dblRotationAngle = CDbl(Tp.Attribute(DEF_ATT_DOOR_ROTATION_ANGLE))
        cNestRect.AddPart gstr_StripExtension(strNestFile), 1, dblRotationAngle, 1, 0
        cNestTrue.AddPart gstr_StripExtension(strNestFile), 1, dblRotationAngle, 1, 0
        Exit For
      End If
    Next
  Next

  Set N = Nesting
  N.DeleteAllNestLists

  '..now do nesting
  With cNestRect

      ' Load the nest list
      Set Nl = N.LoadNestList(.NestListName)
      
      ActiveDrawing.GetExtent 0, 0, 0, 0, dblDrawingExtentY, 0

      dblDrawingExtentY = dblDrawingExtentY + (0.18 * .SheetLength)

      '..create the sheet
      Set NstSheetPath = ActiveDrawing.CreateRectangle(0, dblDrawingExtentY, .SheetWidth, dblDrawingExtentY + .SheetLength)
      '..and the Sheetlist
      Set Sl = N.NewSheetList

      '..add the sheet to the sheetlist
      Set Nsh = Sl.add(NstSheetPath)

      ' Set the sheet properties
      With Nsh
        .MaterialName = clsNest.SheetName
        .Thickness = clsNest.SheetThickness
        .Required = 0
      End With

  End With

  'If ActiveDrawing.ThreeDViews = True Then
  '  ActiveDrawing.ThreeDViews = False
  'End If
    
  '..nest it
  
  intNumSheetsPre = Ni.Sheets.Count
  Set NestErrorsNl = N.Nest(Nl, Sl)
    
  Set Ni = ActiveDrawing.GetNestInformation
  intNumSheetsPost = Ni.Sheets.Count

  '..let's see it, baby!
  'ActiveDrawing.ZoomAll

  ' Test to ensure all components have been nested
  If NestErrorsNl.Count > 0 Then
    For Each Npi In NestErrorsNl
      strNestErrors = strNestErrors & Npi.Required & " x " & Npi.Paths(1).Attribute(DEF_ATT_TYPE_NAME) & " (" & Npi.Paths(1).Attribute(DEF_ATT_PART_WIDTH) & " * " & Npi.Paths(1).Attribute(DEF_ATT_PART_LENGTH) & ")" & Chr(13)
    Next

    If ActiveDrawing.ThreeDViews = True Then ActiveDrawing.ThreeDViews = False

    ActiveDrawing.ZoomAll

    MsgBox Frame.ReadTextFile(strCTX, 500, 127) & Space(3) & Chr(13) & Chr(13) & _
      NestMaterial.MaterialName & Chr(13) & Chr(13) & _
      strNestErrors, vbExclamation, DEF_PROJECT_NAME
    GoTo Controlled_Exit
      
  End If

  If (intNumSheetsPost - intNumSheetsPre) > 1 Then
      
      ' Re-nested result takes more than one sheet
      ' Reject all sheets
      For intSheetCount = (intNumSheetsPre + 1) To intNumSheetsPost
        Set Nsh = Ni.Sheets(intSheetCount)
        m_DeleteNestSheet Nsh
      Next
  
      For Each pthConsGeo In ActiveDrawing.Layers(2).Geometries
        If pthConsGeo.Attribute("LicomUKja_cutline_mark") <> "" Then
          pthConsGeo.Selected = True
        End If
      Next
      ActiveDrawing.DeleteSelected
  
  Else
  
      ' Analyse sheet packing - determine the extent of the left-most toolpath
      Set Nsh = Ni.Sheets(Ni.Sheets.Count)
      
      For Each Npi In Nsh.Parts
        For Each Tp In Npi.Paths
          If Tp.IsToolPath Then
            Tp.GetFeedExtent 0, 0, dblLeftMost, 0
            If dblLeftMost > dblMaxLeftmost Then
              dblMaxLeftmost = dblLeftMost
            End If
          End If
        Next
      Next
      dblLeftmostTPRect = dblMaxLeftmost
      
  End If

  Set N = Nesting
  N.DeleteAllNestLists

  Set Ni = ActiveDrawing.GetNestInformation
  intNumSheetsPre = Ni.Sheets.Count

  '..now do nesting
  With cNestTrue

      ' Load the nest list
      Set Nl = N.LoadNestList(.NestListName)
      
      ActiveDrawing.GetExtent 0, 0, 0, 0, dblDrawingExtentY, 0

      dblDrawingExtentY = dblDrawingExtentY + (0.18 * .SheetLength)

      '..create the sheet
      Set NstSheetPath = ActiveDrawing.CreateRectangle(0, dblDrawingExtentY, .SheetWidth, dblDrawingExtentY + .SheetLength)
      '..and the Sheetlist
      Set Sl = N.NewSheetList

      '..add the sheet to the sheetlist
      Set Nsh = Sl.add(NstSheetPath)

      ' Set the sheet properties
      With Nsh
        .MaterialName = clsNest.SheetName
        .Thickness = clsNest.SheetThickness
        .Required = 0
      End With

  End With

  'If ActiveDrawing.ThreeDViews = True Then
  '  ActiveDrawing.ThreeDViews = False
  'End If
    
  '..nest it
  Set NestErrorsNl = N.Nest(Nl, Sl)
  
  Set Ni = ActiveDrawing.GetNestInformation
  intNumSheetsPost = Ni.Sheets.Count

  '..let's see it, baby!
  'ActiveDrawing.ZoomAll

  ' Test to ensure all components have been nested
  If NestErrorsNl.Count > 0 Then
    For Each Npi In NestErrorsNl
      strNestErrors = strNestErrors & Npi.Required & " x " & Npi.Paths(1).Attribute(DEF_ATT_TYPE_NAME) & " (" & Npi.Paths(1).Attribute(DEF_ATT_PART_WIDTH) & " * " & Npi.Paths(1).Attribute(DEF_ATT_PART_LENGTH) & ")" & Chr(13)
    Next

    If ActiveDrawing.ThreeDViews = True Then ActiveDrawing.ThreeDViews = False

    ActiveDrawing.ZoomAll
  
    MsgBox Frame.ReadTextFile(strCTX, 500, 127) & Space(3) & Chr(13) & Chr(13) & _
      NestMaterial.MaterialName & Chr(13) & Chr(13) & _
      strNestErrors, vbExclamation, DEF_PROJECT_NAME
    GoTo Controlled_Exit
      
  End If

  If (intNumSheetsPost - intNumSheetsPre) > 1 Then
  
      ' Re-nested result takes more than one sheet
      ' Reject all sheets
      For intSheetCount = (intNumSheetsPre + 1) To intNumSheetsPost
        Set Nsh = Ni.Sheets(intSheetCount)
        m_DeleteNestSheet Nsh
      Next
  
  Else
      
      ' Analyse sheet packing - determine the extent of the left-most toolpath
      Set Nsh = Ni.Sheets(Ni.Sheets.Count)
      dblMaxLeftmost = 0
      
      For Each Npi In Nsh.Parts
        For Each Tp In Npi.Paths
          If Tp.IsToolPath Then
            Tp.GetFeedExtent 0, 0, dblLeftMost, 0
            If dblLeftMost > dblMaxLeftmost Then
              dblMaxLeftmost = dblLeftMost
            End If
          End If
        Next
      Next
      dblLeftmostTPTrue = dblMaxLeftmost
      
  End If
  
  Set Ni = ActiveDrawing.GetNestInformation
  
  If dblLeftmostTPTrue <> 0 And dblLeftmostTPRect <> 0 Then
  
      ' Both Algorithms have succeeded
      ' Compare the two - reject original nest
      Set Nsh = Ni.Sheets(strLastSheetName)
      m_DeleteNestSheet Nsh
  
      ' Use the rectangular nesting packing unless the true shape is better
      If dblLeftmostTPTrue < dblLeftmostTPRect Then
        ' Reject the rectangular packed sheet
        m_DeleteNestSheet Ni.Sheets(Ni.Sheets.Count - 1)
        For Each pthConsGeo In ActiveDrawing.Layers(2).Geometries
          If pthConsGeo.Attribute("LicomUKja_cutline_mark") <> "" Then
            pthConsGeo.Selected = True
          End If
        Next
        ActiveDrawing.DeleteSelected
      Else
        ' Reject the true shape packed sheet
        m_DeleteNestSheet Ni.Sheets(Ni.Sheets.Count)
      End If
  
  ElseIf dblLeftmostTPTrue = 0 And dblLeftmostTPRect <> 0 Then
    
    ' True shape failed (>1 sheet) but rectangular worked
    ' Use rectangular - reject original nest
    Set Nsh = Ni.Sheets(strLastSheetName)
    m_DeleteNestSheet Nsh
  
  ElseIf dblLeftmostTPTrue <> 0 And dblLeftmostTPRect = 0 Then
  
    ' True shape worked but rectangular failed
    ' Use true shape - reject original nest
    Set Nsh = Ni.Sheets(strLastSheetName)
    m_DeleteNestSheet Nsh
    
  ElseIf dblLeftmostTPTrue = 0 And dblLeftmostTPRect = 0 Then
    
    ' Both will have already been deleted
    ' Do nothing - leave original nest
    
  End If


Controlled_Exit:

  Set N = Nothing
  Set Ni = Nothing
  Set Nl = Nothing
  Set NestErrorsNl = Nothing
  Set Sl = Nothing
  Set Nsh = Nothing
  Set NstSheetPath = Nothing
  Set Npi = Nothing
  Set Tp = Nothing
  Set pthConsGeo = Nothing
  Set cNestRect = Nothing
  Set cNestTrue = Nothing

Exit Sub

m_PackFinalNestSheetToLHS_Error:

  MsgBox "Error in m_PackFinalNestSheetToLHS:" & Chr(13) & Err.Description, vbExclamation, DEF_PROJECT_NAME
  
  Resume Controlled_Exit

End Sub


Public Sub m_PopulateNestingZones()
  Dim rstZone As ADODB.Recordset
'
  On Error GoTo m_PopulateNestingZones_Error
  
  Set colManualNestZones = Nothing
  Set colAutoNestZones = Nothing
  
  Set pthsNestZones = ActiveDrawing.CreatePathCollection
  
  Set rstZone = grst_GetNestingZones(adoorNestZone_Manual)
  
  If Not rstZone Is Nothing Then
    m_PopulateNestZoneCollection rstZone, colManualNestZones
  End If
  
  Set rstZone = grst_GetNestingZones(adoorNestZone_Automatic)
  
  If Not rstZone Is Nothing Then
    m_PopulateNestZoneCollection rstZone, colAutoNestZones
  End If
    
    
m_PopulateNestingZones_Error:
  
  On Error Resume Next
  
  If rstZone.State = adStateOpen Then
    rstZone.Close
  End If

  Set rstZone = Nothing

End Sub


Private Sub m_PopulateNestZoneCollection(rstNestingZones As ADODB.Recordset, TargetCollection As Collection)
  
  Dim NestZone As CNestingZone
  
  On Error GoTo m_PopulateNestZoneCollection_Err
  
  Do While Not rstNestingZones.EOF
    Set NestZone = New CNestingZone
    With NestZone
      .ZoneID = gvar_CheckNull(rstNestingZones.Fields!ZoneID)
      .ZoneName = gvar_CheckNull(rstNestingZones.Fields!ZoneName)
      .ZoneNumber = gvar_CheckNull(rstNestingZones.Fields!ZoneNumber)
      .ZoneStartX = gvar_CheckNull(rstNestingZones.Fields!ZoneStartPointX)
      .ZoneStartY = gvar_CheckNull(rstNestingZones.Fields!ZoneStartPointY)
      .ZoneWidth = gvar_CheckNull(rstNestingZones.Fields!ZoneWidth)
      .ZoneHeight = gvar_CheckNull(rstNestingZones.Fields!ZoneHeight)
      .ZonePartDimension = gvar_CheckNull(rstNestingZones.Fields!ZonePartDimension)
      .ZonePartArea = gvar_CheckNull(rstNestingZones.Fields!ZonePartArea)
      ' TFS#65726
      .ZoneAssignMethod = gvar_CheckNull(rstNestingZones.Fields!ZoneAssignMethod)
    
      TargetCollection.add NestZone, "k" & .ZoneID
    
    End With
    
    rstNestingZones.MoveNext
  Loop
  
m_PopulateNestZoneCollection_Err:

  If rstNestingZones.State = adStateOpen Then
    rstNestingZones.Close
  End If
  
End Sub


Private Sub m_PopulateRulesCollection(CurrentRouter As CRouter)

  Dim rstRules          As ADODB.Recordset
  Dim Rule              As CRule

  On Error GoTo m_PopulateRulesCollection_Error
  
  ' Reset rules collection
  Set colRouterRules = Nothing
    
  Set rstRules = grst_GetRulesForRouter(gstr_StripLicomDatPath(CurrentRouter.PostProcessor))
  
  ' Test to see if there are any rules to apply
  If rstRules Is Nothing Then
    GoTo Controlled_Exit
  End If
  
  rstRules.MoveFirst
    
  Do While Not rstRules.EOF
    
    Set Rule = New CRule
    
    With Rule
      
      .RuleID = gvar_CheckNull(rstRules.Fields!RuleID)
      .OperatorID = gvar_CheckNull(rstRules.Fields!OperatorID)
      .RuleName = gvar_CheckNull(rstRules.Fields!RuleName)
      .TestVariableID = gvar_CheckNull(rstRules.Fields!TestVariableName)
      .TestVariableValue = gvar_CheckNull(rstRules.Fields!TestVariableValue)
      .ResultRouterPost = gvar_CheckNull(rstRules.Fields!ResultRouterPost)
      .ResultRouter = gvar_CheckNull(rstRules.Fields!ResultRouter)
      .RuleText = gvar_CheckNull(rstRules.Fields!RuleText)
                  
    End With
    
    rstRules.MoveNext
  
    colRouterRules.add Rule
  
  Loop
  
  rstRules.Close

Controlled_Exit:

  Set Rule = Nothing
  Set rstRules = Nothing

Exit Sub

m_PopulateRulesCollection_Error:

  MsgBox Err.Description, vbExclamation, "m_PopulateRulesCollection"
  If (Err.Number <> 0) Then WriteError Err, True, "m_PopulateRulesCollection"
  
  Resume Controlled_Exit

End Sub

Private Function mbln_RulesApply(Door As CDoor) As Boolean

  Dim Rule            As CRule
  Dim RuleTestResult  As AdoorRuleTestResult
'
  On Error GoTo mbln_RulesApply_Error
  
  ' Assume success
  mbln_RulesApply = True
    
  For Each Rule In colRouterRules
  
    With Rule
      
      RuleTestResult = m_RuleTestValid(Door, .TestVariableID, .OperatorID, .TestVariableValue)
        
      Select Case RuleTestResult
          Case adoorRuleTestResultTrue
              ' The rule condition has been met
              Select Case .ResultRouter
                Case adoorRULE_RESULT_ROTATE_LOCK
                  Door.RotationMethod = adoorPART_ROTATION_LOCKX
                Case adoorRULE_RESULT_LOCK_ONLY
                  Door.RotationMethod = adoorPART_ROTATION_LOCKY
                Case adoorRULE_RESULT_FREE_ROTATE
                  Door.RotationMethod = adoorPART_ROTATION_FREE
              End Select
          
          Case adoorRuleTestResultFalse
          ' The rule test criteria was not met - nothing to do
              
          Case adoorRuleTestResultError
            ' An error occurred
            MsgBox Frame.ReadTextFile(strCTX, 500, 138) & Chr(13) & _
              .RuleText, vbExclamation, DEF_PROJECT_NAME
            mbln_RulesApply = False
            
      End Select
  
    
    End With
  
  Next

Controlled_Exit:

Exit Function

mbln_RulesApply_Error:

  mbln_RulesApply = False
  MsgBox Err.Description, vbExclamation, "mbln_RulesApply"
  If (Err.Number <> 0) Then WriteError Err, True, "mbln_RulesApply"
  
  Resume Controlled_Exit

End Function

Private Function m_RuleTestValid(Door As CDoor, TestItem As AdoorRuleTestType, TestOperatorID As Long, TestValue As Double) As AdoorRuleTestResult

  Dim dblValue As Double
'
  On Error GoTo mbln_RuleTestValid_Error
  
  If TestItem = adoorRULE_TEST_TYPE_WIDTH Then
    dblValue = Door.Width
  ElseIf TestItem = adoorRULE_TEST_TYPE_HEIGHT Then
    dblValue = Door.Length
  End If
  
  ' See which operator we are using
  Select Case TestOperatorID
    Case adoorRuleOperatorEqualTo
      If dblValue = TestValue Then
        m_RuleTestValid = adoorRuleTestResultTrue
      Else
        m_RuleTestValid = adoorRuleTestResultFalse
      End If
    
    Case adoorRuleOperatorGreaterThan
      If dblValue > TestValue Then
        m_RuleTestValid = adoorRuleTestResultTrue
      Else
        m_RuleTestValid = adoorRuleTestResultFalse
      End If
    
    Case adoorRuleOperatorLessThan
      If dblValue < TestValue Then
        m_RuleTestValid = adoorRuleTestResultTrue
      Else
        m_RuleTestValid = adoorRuleTestResultFalse
      End If
    
    Case adoorRuleOperatorGreaterThanOrEqualTo
      If dblValue >= TestValue Then
        m_RuleTestValid = adoorRuleTestResultTrue
      Else
        m_RuleTestValid = adoorRuleTestResultFalse
      End If
    
    Case adoorRuleOperatorLessThanOrEqualTo
      If dblValue <= TestValue Then
        m_RuleTestValid = adoorRuleTestResultTrue
      Else
        m_RuleTestValid = adoorRuleTestResultFalse
      End If
    
    Case adoorRuleOperatorNotEqualTo
      If dblValue <> TestValue Then
        m_RuleTestValid = adoorRuleTestResultTrue
      Else
        m_RuleTestValid = adoorRuleTestResultFalse
      End If
  
    Case Else
      ' Should not get here
      m_RuleTestValid = adoorRuleTestResultError
  
  End Select

Controlled_Exit:

Exit Function

mbln_RuleTestValid_Error:

  m_RuleTestValid = adoorRuleTestResultError
  MsgBox Err.Description, vbExclamation, "m_RuleTestValid"
  If (Err.Number <> 0) Then WriteError Err, True, "m_RuleTestValid"
  
  Resume Controlled_Exit

End Function


Private Sub m_ToolSorting()
  
  ' Test to see if tool ordering has been used
  If clsOptions.ToolOrderingUsed Then
    
    SuppressUpdateRapids True
    
    If Not mbln_ToolOrdering Then
      MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 137), vbExclamation, DEF_PROJECT_NAME
    End If
    
    SuppressUpdateRapids False
  End If
  
End Sub

Private Function mbln_ToolOrdering() As Boolean
  Dim rstToolOrder            As ADODB.Recordset
  Dim lngOpCounter            As Long
  Dim lngOpInsertPoint        As Long
  Dim strToolOrderName        As String
  Dim strOpToolName           As String
  Dim lngOpToolNumber         As Long
  Dim lngOpToolOffset         As Long
  Dim lngToolNumber           As Long
  Dim lngToolOffset           As Long
  Dim NestInfo                As NestInformation
  Dim Ns                      As NestSheet
  Dim colSheetID              As New Collection
  Dim iNumSheets              As Integer
  Dim i                       As Integer
  Dim lFirstOp                As Long
  Dim lLastOp                 As Long
  Dim Op                      As Operation
  Dim Ops                     As Operations

  On Error GoTo EH:
  ' Assume success
  mbln_ToolOrdering = True
  
  ' Only re-order if there are 2 or more operations
  If ActiveDrawing.Operations.Count < 2 Then Exit Function
  
  Frame.ShowProgressBox DEF_PROJECT_NAME, "Ordering Toolpaths"
  
  ' Get the tool order from the database
  Set rstToolOrder = grst_GetToolOrder
  
  ' Test to see if a tool order has been defined
  If rstToolOrder Is Nothing Then
    mbln_ToolOrdering = False
    Exit Function
  End If
  
  ' Put all of the sheet ID's into a collection
  Set NestInfo = ActiveDrawing.GetNestInformation
  For Each Ns In NestInfo.Sheets
    ' Do not add reverse sheets to the collection
    If UCase(Right(Ns.Name, 4)) <> "_REV" Then
      'colSheetID.Add DetermineSheetID(Ns.Name)
      colSheetID.add Ns.Name
    End If
  Next
  
  App.DisableUndo = True
  
  iNumSheets = colSheetID.Count
  
  Set Ops = ActiveDrawing.Operations
  
  For i = 1 To iNumSheets
    Frame.SetProgressText Frame.ReadTextFile(strCTX, 410, 4) & Space$(1) & colSheetID(i)
    GetSheetFirstAndLastOperationToolOrdering colSheetID(i), lFirstOp, lLastOp
    
    ' Move to the first record in the recordset
    rstToolOrder.MoveFirst
      
    lngOpInsertPoint = lFirstOp
    
    Do While Not rstToolOrder.EOF
    
      ' Get the Tool details for this recordset tool
      strToolOrderName = rstToolOrder!ToolName
      lngToolNumber = rstToolOrder!ToolNumber
      lngToolOffset = rstToolOrder!ToolOffsetNumber
      
      ' Loop through the operations
      lngOpCounter = lFirstOp
      
      Do
        ' Test to see if the tool name, number and offset of this operation is the same as the tool
        ' details we are looking for from the database recordset
        'strOpToolName = ActiveDrawing.Operations(lngOpCounter).Tool.Name
        'lngOpToolNumber = ActiveDrawing.Operations(lngOpCounter).Tool.Number
        'lngOpToolOffset = ActiveDrawing.Operations(lngOpCounter).Tool.OffsetNumber
        
        ' TFS #57996 - removal of calls to ActiveDrawing.Operations because of large amount of memory used
        Set Op = Nothing
        Set Op = Ops(lngOpCounter)
        
        strOpToolName = Op.Tool.Name
        lngOpToolNumber = Op.Tool.Number
        lngOpToolOffset = Op.Tool.OffsetNumber
        
        If UCase(strOpToolName) = UCase(strToolOrderName) And _
          (lngOpToolNumber = lngToolNumber) And _
          (lngOpToolOffset = lngToolOffset) Then
          
          ' The tool details match - move this operation to the current
          ' insertion point on the operations list
          ActiveDrawing.Operations.Renumber lngOpCounter, lngOpInsertPoint, acamOpINSERT_IN_FRONT
          Set Ops = ActiveDrawing.Operations
          lngOpInsertPoint = lngOpInsertPoint + 1
        End If
        lngOpCounter = lngOpCounter + 1
        
      Loop Until lngOpCounter > lLastOp
    
      ' Move to the next record in the recordset
      rstToolOrder.MoveNext
  
    Loop
      
  Next
    
  App.DisableUndo = False
  
  ' Close the recordset
  rstToolOrder.Close
  
Controlled_Exit:

  Set Op = Nothing
  Set Ops = Nothing

  ' Close the progress box
  Frame.CloseProgressBox
  
Exit Function

EH:

  mbln_ToolOrdering = False
  ' Close the recordset if open
  If rstToolOrder.State = adStateOpen Then rstToolOrder.Close
  MsgBox Err.Description, vbExclamation, "mbln_ToolOrdering"
  If (Err.Number <> 0) Then WriteError Err, True, "mbln_ToolOrdering"
  
  Resume Controlled_Exit

End Function


Public Function gbln_Make_Master_Press() As Boolean
    
    Dim Press               As CPress
    Dim PressColour         As CPressColour
    Dim PressThickness      As CPressThickness
    Dim PressDoor           As CDoor
    Dim blnNestListStarted  As Boolean
    Dim strCTX              As String
    Dim strProgress         As String
    Dim lngPartNumber       As Long
    Dim colDOORANC          As New Collection
    Dim bPress              As Boolean
    
    strCTX = clsOptions.CTXFile
    
    blnNestListStarted = False
    bPress = True
    
    For Each Press In colPressData
      For Each PressColour In Press.colPressColours
      
        '..need new drawing
        App.New

        '..set the progress text
        strProgress = Frame.ReadTextFile(strCTX, 300, 5) & Space(1) & _
                            PressColour.ColourName

        Frame.SetProgressText strProgress
        DoEvents
      
        For Each PressThickness In PressColour.colPressThicknesses
          
          For Each PressDoor In PressThickness.colPressComponents
            
            '..create the nest list header
            If Not blnNestListStarted Then
                Set clsNest = New CNest
                clsNest.StartNestListPress PressDoor.JobName, Press
                clsNest.SheetThickness = PressThickness.PressThickness
                blnNestListStarted = True
            End If
          
            '..make the part
            If Not mbln_ProcessPart(PressDoor, lngPartNumber, colDOORANC, bPress, False) Then

                '..had an issue
                SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_COMPLETENEST, 0

                '..let the user know that something wasn't right
                MsgBox Frame.ReadTextFile(strCTX, 600, 26) & Space(3), vbInformation, DEF_PROJECT_NAME

                App.New
                GoTo Controlled_Exit

            End If
          
          Next
          
          ' Test to see if we are nesting for the press by material thickness
          If clsOptions.GroupByMaterialThickness Then
            ' Nest this colour and thickness now
            If blnNestListStarted Then
                '..start clean so we can nest
                App.New
                '..if no parts have been created then bail
                If CBool(lngPartNumber) Then
                    m_DoNestingPress Press, PressColour, PressThickness.PressThickness
                End If
                blnNestListStarted = False
            End If
          End If
        
        ' Get the next thickness
        Next
        
        ' Test to see if all thicknesses should be grouped together
        If Not clsOptions.GroupByMaterialThickness Then
          ' Nest all thicknesses of this colour now
          If blnNestListStarted Then
              '..start clean so we can nest
              App.New
              '..if no parts have been created then bail
              If CBool(lngPartNumber) Then
                  m_DoNestingPress Press, PressColour
              End If
              blnNestListStarted = False
          End If
        End If
        
      ' Get the next Foil colour
      Next
    
    ' Get the next press machine tool
    Next
    
    
'
'                    '..Remove all nestlists
'                    Set N = Nesting
'                    N.DeleteAllNestLists
'
'                    '..now do nesting
'                    With clsNest
'
'                        ' Load the nest list
'                        Set Nl = N.LoadNestList(.NestListName)
'
'                        ' If the user has selected to order the nested components so that
'                        ' the biggest parts are numbered first, update the nestlist
'                        If .NumberComponentsBySize Then
'                          Nl.OrderParts
'                        End If
'
'                        '..create the sheet
'                        Set nstSheet = drw.CreateRectangle(0, 0, .SheetWidth, .SheetLength)
'                        '..and the Sheetlist
'                        Set Sl = N.NewSheetList
'
'                        '..add the sheet to the sheetlist
'                        Set Nsh = Sl.Add(nstSheet)
'
'                        ' Set the sheet properties
'                        With Nsh
'                          .MaterialName = clsNest.SheetName
'                          .Thickness = clsNest.SheetThickness
'                          .Required = 0
'                        End With
'
'                    End With
'
'                    ' Set a global attribute to indicate the current material being nested
'                    ActiveDrawing.Attribute(DEF_ATT_MATERIAL_NAME) = CStr(rstMaterials.Fields!Name)
'
'                    ' Set an attribute to indicate that Twin Head Nesting can begin
'                    ' on the next NestingComplete event
'                    ActiveDrawing.Attribute(DEF_ATT_CAN_START_TWIN_HEAD) = "1"
'
'                    ' Clear the twin head nesting registry flag
'                    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_TWINHEAD, "0"
'
'                    ' Clear the reverse side nesting registry flag
'                    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_REVERSENEST, "0"
'
'                    '..nest it
'                    Set NestErrorsNl = N.Nest(Nl, Sl)
'
'                    '..let's see it, baby!
'                    drw.ZoomAll
'
'                    ' Test to ensure all components have been nested
'                    If NestErrorsNl.Count > 0 Then
'                      For Each Np In NestErrorsNl
'                        strNestErrors = strNestErrors & Np.Required & " x " & Np.Paths(1).Attribute(DEF_ATT_TYPE_NAME) & " (" & Np.Paths(1).Attribute(DEF_ATT_PART_WIDTH) & " * " & Np.Paths(1).Attribute(DEF_ATT_PART_LENGTH) & ")" & Chr(13)
'                      Next
'
'                      MsgBox Fr.ReadTextFile(sCTX, 500, 127) & Space(3) & Chr(13) & Chr(13) & _
'                        strPromptMaterial & Chr(13) & Chr(13) & _
'                        strNestErrors, vbExclamation, DEF_PROJECT_NAME
'                      GoTo Controlled_Exit
'                    Else
'                      If Not mbln_CompareNestToOrder(strOrder, rstMaterials.Fields!Name) Then
'                        MsgBox Fr.ReadTextFile(sCTX, 500, 117) & Space(3) & Chr(13) & _
'                          Fr.ReadTextFile(sCTX, 500, 118), vbExclamation, DEF_PROJECT_NAME
'                          GoTo Controlled_Exit
'                      End If
'                    End If
'
'                    '..save nest list?
'                    If Not clsOptions.SaveAllNestANL Then
'
'                        With FSO
'                            If .FileExists(clsNest.NestListName) Then .DeleteFile (clsNest.NestListName)
'                        End With
'
'                    End If
'
'                    ' Remove all nestlists
'                    N.DeleteAllNestLists
'
'                    '..insert report data
'                        If Not clsOptions.DisableReports Then
'                            Call m_InsertReportData(drw, Fr, sCTX, udtDTD, DrwMaterial.Name, rstMaterials.AbsolutePosition, lngSheetNumber)
'                        End If
'
'                    '..save the program
'                    With clsOptions
'
'                        '..set the post variables
'                        Call m_SetPostVariables(clsNest.SheetWidth, clsNest.SheetLength, _
'                                                clsNest.SheetThickness, _
'                                                ((lngNestNumber + 1) * 1000) + 1, .PromptForProgram)
'                                                 'lngNestNumber + 1001)
'
'                        strNestANCName = mstr_CompileNestOrNCFilename(strJob, clsNest.SheetName, rstMaterials.AbsolutePosition)
'                        'strNestANCName = DEF_NEST_PREFIX & strJob & DEF_UNDERSCORE & clsNest.SheetName
'
'                        strNestANCFullName = .PathToRoot & strNestANCName & "." & clsOptions.NCFileExtension
'
'                        ' Test for MPR output
'                        If clsOptions.OutputMPR Then
'                          With clsNest
'                            m_MprSave CStr(.SheetLength), CStr(.SheetWidth), CStr(clsNest.SheetThickness), "0", "0", "0", "0", "0", "0", "0", "", "0", clsOptions.PathToRoot & strNestANCName & DEF_EXTENSION_MPR
'                          End With
'
'                        Else
'                          '..output nc to file
'                          drw.OutputNC strNestANCFullName, acamOutNcFILE, False
'                        End If
'
'                        '..up the part count (nested files)
'                        lngNestNumber = lngNestNumber + 1
'
'                        '..collect the nesting information
'                        Set nstInfo = drw.GetNestInformation
'
'                        '..split programs if more than 1 sheet and option enabled
'                        If ((nstInfo.Sheets.Count > 1) And (.SplitNestedPrograms)) Then
'
'                            Call m_SplitNestANC(lngOrder, nstInfo, strNestANCFullName, strNestANCName, colNestANC)
'
'                        Else
'
'                            If Not .OutputMPR Then
'                              '..added nested anc file to array
'                              colNestANC.Add strNestANCFullName
'                            End If
'
'                            '..update the nest anc path in the database
'                            If Not .DisableReports Then
'                                Call m_UpdateDBNestPaths(lngOrder, DEF_RST_RPT_NESTANC, DEF_RST_RPT_PATHTONESTANC, strNestANCFullName, strNestANCName, False)
'                            End If
'
'                        End If
'
'                        '..are we saving it
'                        If .SaveAllNestARD Then
'
'                            '..assign the nested ard file name
'                            strNestARD = mstr_CompileNestOrNCFilename(strJob, clsNest.SheetName, rstMaterials.AbsolutePosition) & DEF_EXTENSION_ARD
'                            'strNestARD = DEF_NEST_PREFIX & strJob & DEF_UNDERSCORE & clsNest.SheetName & DEF_EXTENSION_ARD
'
'                            '..save to second location?  this is done first so that default path is used for reports    '..07.21.02 - rg
'                            If .OutputToSecondLocation Then
'                                    drw.SaveAs .PathToRootSecond & strNestARD
'
'                                    ' 11/21/05 - rg
'                                    '
'                                    Call drw.AddToRecentFileList
'
'                            End If
'
'                            '..save the nested drawing
'                            drw.SaveAs .PathToRoot & strNestARD
'
'                            ' 11/21/05 - rg
'                            '
'                            Call drw.AddToRecentFileList
'
'                            '..update the nest ard path in the database
'                            If Not .DisableReports Then
'                                Call m_UpdateDBNestPaths(lngOrder, DEF_RST_RPT_NESTARD, DEF_RST_RPT_PATHTONESTARD, drw.FullName, drw.Name & DEF_EXTENSION_ARD, False)
'                            End If
'
'
'
'                        End If
'
'                        ' Test to see if the file naming convention is using a sequential order number
'                        ' rather than the order name
'                        If Not .UseOrderName Then
'                          ' increment the sequential number
'                          If Not gbln_IncrementNCSeqNum Then
'                            MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 124), vbExclamation, DEF_PROJECT_NAME
'                          End If
'                        End If
'
'                        '..save single drawings when nesting?
'                        If Not udtDTD.ByPassNest Then
'                            On Error Resume Next
'                            If Not .SaveAllDoorARD Then FSO.DeleteFile .PathToRoot & strJob & DEF_WILDCARD & DEF_EXTENSION_ARD, True
'                            On Error GoTo g_Make_Master_Error
'                        End If
'
'                    End With
'
'                End If
'
'            End If
'
'        End If
'
'        '..get the next material
'        rstMaterials.MoveNext
'
'    Loop '..next material
'
Controlled_Exit:

End Function

Public Function gstr_GenerateNestedSheetBarcode() As String
    Dim cINI          As CIniFile
    Dim sini          As String
    Dim strBarcodeID  As String
    Dim lngBarcodeID  As Long
    Dim FSO           As New Scripting.FileSystemObject
    
    ' 08 jun 11 TFS#44783
    '
    'sINI = Frame.PathOfThisAddin & "\Barcode.ini"
    Call g_MoveBarcodeIni(sini)
    
    Set cINI = New CIniFile

    '..see if .ini file exists
    If FSO.FileExists(sini) Then
    
        With cINI
            .Init sini
            strBarcodeID = .GetStringValue("SheetID", "NextSheetID")
            gstr_GenerateNestedSheetBarcode = strBarcodeID
            
            lngBarcodeID = CLng(strBarcodeID)
            lngBarcodeID = lngBarcodeID + 1
            strBarcodeID = CStr(lngBarcodeID)
            
            .SaveStringValue "SheetID", "NextSheetID", strBarcodeID
        
        End With
    End If

    Set cINI = Nothing

End Function

Private Sub m_DoNestingPress(Press As CPress, PressColour As CPressColour, Optional PressThickness As Double)
    
    Dim N                       As Nesting
    Dim Ni                      As NestInformation
    Dim Nl                      As Nestlist
    Dim NestErrorsNl            As Nestlist
    Dim Sl                      As Sheetlist
    Dim Nsh                     As NestSheet
    Dim Np                      As NestPart
    Dim NstSheet                As Path
    Dim strNestErrors           As String
    Dim FSO                     As New Scripting.FileSystemObject
    Dim strPressFileName        As String
    Dim strNestARD              As String

    '..Remove all nestlists
    Set N = Nesting
    N.DeleteAllNestLists

    '..now do nesting
    With clsNest

        ' Load the nest list
        Set Nl = N.LoadNestList(.NestListName)

        ' If the user has selected to order the nested components so that
        ' the biggest parts are numbered first, update the nestlist
        If .NumberComponentsBySize Then
          Nl.OrderParts
        End If

        '..create the sheet
        Set NstSheet = ActiveDrawing.CreateRectangle(0, 0, .SheetWidth, .SheetLength)
        '..and the Sheetlist
        Set Sl = N.NewSheetList

        '..add the sheet to the sheetlist
        Set Nsh = Sl.add(NstSheet)

        ' Set the sheet properties
        With Nsh
          .MaterialName = clsNest.SheetName
          .Thickness = clsNest.SheetThickness
          .Required = 0
        End With

    End With

    If ActiveDrawing.ThreeDViews = True Then
      ActiveDrawing.ThreeDViews = False
    End If
    
    '..nest it
    Set NestErrorsNl = N.Nest(Nl, Sl)

    '..let's see it, baby!
    ActiveDrawing.ZoomAll

    ' Test to ensure all components have been nested
    If NestErrorsNl.Count > 0 Then
      For Each Np In NestErrorsNl
        strNestErrors = strNestErrors & Np.Required & " x " & Np.Paths(1).Attribute(DEF_ATT_TYPE_NAME) & " (" & Np.Paths(1).Attribute(DEF_ATT_PART_WIDTH) & " * " & Np.Paths(1).Attribute(DEF_ATT_PART_LENGTH) & ")" & Chr(13)
      Next

      MsgBox Frame.ReadTextFile(strCTX, 500, 127) & Space(3) & Chr(13) & Chr(13) & _
        strNestErrors, vbExclamation, DEF_PROJECT_NAME
      GoTo Controlled_Exit
    Else
      If Not mbln_ComparePressNestToOrder(Press, PressColour, PressThickness) Then
        MsgBox Frame.ReadTextFile(strCTX, 500, 117) & Space(3) & Chr(13) & _
          Frame.ReadTextFile(strCTX, 500, 118), vbExclamation, DEF_PROJECT_NAME
          GoTo Controlled_Exit
      End If
    End If

    Set Ni = ActiveDrawing.GetNestInformation
    
    For Each Nsh In Ni.Sheets
      Nsh.Path.Attribute(DEF_ATT_PRESS_NAME) = Press.PressName
      Nsh.Path.Attribute(DEF_ATT_FOIL_COLOUR) = PressColour.ColourName
      Nsh.Path.Attribute(DEF_ATT_SHEET_THICKNESS) = Nsh.Thickness
    Next
        
    With clsOptions
    
        '..save nest list?
        If Not .SaveAllNestANL Then
    
            With FSO
                If .FileExists(clsNest.NestListName) Then .DeleteFile (clsNest.NestListName)
            End With
    
        End If
        
        If PressThickness <> 0 Then
          strPressFileName = Press.PressName & DEF_UNDERSCORE & PressColour.ColourName & PressThickness
        Else
          strPressFileName = Press.PressName & DEF_UNDERSCORE & PressColour.ColourName
        End If
        
        '..save to second location?  this is done first so that default path is used for reports    '..07.21.02 - rg
        If .OutputToSecondLocation Then
                
            If .OutputSecondNestedDrawings Then
            
                '..assign the nested ard file name
                strNestARD = mstr_CompileNestOrNCFilenamePress(strPressFileName) & DEF_EXTENSION_ARD
                
                '..save the nested drawing
                If clsOptions.OutputSecondCreateSubFolder Then
                  ActiveDrawing.SaveAs gstr_EnsureBackslash(.PathToRootSecond) & gstr_JobName & "\" & strNestARD
                Else
                  ActiveDrawing.SaveAs gstr_EnsureBackslash(.PathToRootSecond) & strNestARD
                End If

            End If
        
        End If
        
        '..are we saving it
        If .SaveAllNestARD Then
            
            '..assign the nested ard file name
            strNestARD = mstr_CompileNestOrNCFilenamePress(strPressFileName) & DEF_EXTENSION_ARD
    
            '..save the nested drawing
            If clsOptions.OutputResultsSubFolder Then
              ActiveDrawing.SaveAs gstr_EnsureBackslash(.PathToRoot) & gstr_JobName & "\" & strNestARD
            Else
              ActiveDrawing.SaveAs gstr_EnsureBackslash(.PathToRoot) & strNestARD
            End If
            
            ' 11/21/05 - rg
            '
            Call ActiveDrawing.AddToRecentFileList
    
        End If
    
        If Not clsOptions.DisableReports Then
            Call m_InsertReportDataPress(Press, PressColour, PressThickness)
        End If
    
    End With


Controlled_Exit:

End Sub

Private Sub m_DoNestingRouter(Router As CRouter, Material As CMaterial, colNestANC As Collection)
    
    Dim N                       As Nesting
    Dim Nl                      As Nestlist
    Dim NestErrorsNl            As Nestlist
    Dim Sl                      As Sheetlist
    Dim Nsh                     As NestSheet
    Dim Np                      As NestPart
    Dim Ni                      As NestInformation
    Dim NstSheet                As Path
    Dim strNestErrors           As String
    Dim FSO                     As New Scripting.FileSystemObject
    Dim TwinHeadNestingFlag     As String
    Dim ReverseSideNestingFlag  As String
    Dim lngNestNumber           As Long
    Dim strNestARD              As String
    Dim strNestANCName          As String
    Dim strNestANCFullName      As String
    Dim nstInfo                 As NestInformation
    Dim strSave                 As String
    Dim NestZone                As CNestingZone
    Dim pthZone                 As Path
    
    '..Remove all nestlists
    Set N = Nesting
    N.DeleteAllNestLists

    '..now do nesting
    With clsNest

        ' Load the nest list
        Set Nl = N.LoadNestList(.NestListName)

        ' If the user has selected to order the nested components so that
        ' the biggest parts are numbered first, update the nestlist
        If .NumberComponentsBySize Then
          Nl.OrderParts
        End If

        '..create the sheet
        Set NstSheet = ActiveDrawing.CreateRectangle(0, 0, .SheetWidth, .SheetLength)
        '..and the Sheetlist
        Set Sl = N.NewSheetList
        
''        EditMark
        Dim dll As Object
        Set dll = CreateObject("StdAlpha.ShareClass")
        dll.ScrapNesting clsNest.SheetName, clsNest.SheetThickness, .SheetWidth, Sl

        '..add the sheet to the sheetlist
        Set Nsh = Sl.add(NstSheet)

        ' Set the sheet properties
        With Nsh
          .MaterialName = clsNest.SheetName
          .Thickness = clsNest.SheetThickness
          .Required = 0
        End With
        
        ' Add any nesting Zones
        m_DrawNestZones Material, colAutoNestZones
        m_DrawNestZones Material, colManualNestZones

    End With

    ' Set an attribute to indicate that Twin Head Nesting can begin
    ' on the next NestingComplete event
    ActiveDrawing.Attribute(DEF_ATT_CAN_START_TWIN_HEAD) = "1"

    ActiveDrawing.Attribute(DEF_ATT_TWINHEAD_ORDERSTRING) = gstr_JobIDs

    ' Clear the twin head nesting registry flag
    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_TWINHEAD, "0"

    ' Clear the reverse side nesting registry flag
    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_REVERSENEST, "0"

    ' Turn off SuppressRapids before nesting
    ' (In case Twin head nesting or other macros are set to run on the NestComplete event)
    SuppressUpdateRapids False
    
    If ActiveDrawing.ThreeDViews = True Then
      ActiveDrawing.ThreeDViews = False
    End If
    
    '..nest it
    Set NestErrorsNl = N.Nest(Nl, Sl)

    '..let's see it, baby!
    ActiveDrawing.ZoomAll

    ' Test to ensure all components have been nested
    If NestErrorsNl.Count > 0 Then
      For Each Np In NestErrorsNl
        strNestErrors = strNestErrors & Np.Required & " x " & Np.Paths(1).Attribute(DEF_ATT_TYPE_NAME) & " (" & Np.Paths(1).Attribute(DEF_ATT_PART_WIDTH) & " * " & Np.Paths(1).Attribute(DEF_ATT_PART_LENGTH) & ")" & Chr(13)
      Next

      MsgBox Frame.ReadTextFile(strCTX, 500, 127) & Space(3) & Chr(13) & Chr(13) & _
        Material.MaterialName & Chr(13) & Chr(13) & _
        strNestErrors, vbExclamation, DEF_PROJECT_NAME
      GoTo Controlled_Exit
    Else
      If Not mbln_CompareRouterNestToOrder(Router, Material.MaterialName) Then
        MsgBox Frame.ReadTextFile(strCTX, 500, 117) & Space(3) & Chr(13) & _
          Frame.ReadTextFile(strCTX, 500, 118), vbExclamation, DEF_PROJECT_NAME
          GoTo Controlled_Exit
      End If
    End If

    '..save nest list?
    If Not clsOptions.SaveAllNestANL Then

        With FSO
            If .FileExists(clsNest.NestListName) Then .DeleteFile (clsNest.NestListName)
        End With

    End If

    '..Pack final sheet to LHS if required
    If clsNest.PackFinalSheetToLHS Then
    
        m_PackFinalNestSheetToLHS Material
    
    End If
    
    '..insert scrap filler parts?
    If clsNest.InsertFiller Then

        ' Disable Twin Head Nesting (if appropriate)
        ActiveDrawing.Attribute(DEF_ATT_CAN_START_TWIN_HEAD) = "0"
        
        If Not mbln_InsertFillerParts(Material) Then
    
            '..let them know if failed
            MsgBox Frame.ReadTextFile(strCTX, 500, 56) & Space(3) & vbCrLf & _
                   clsNest.InsertFillerFile & Space(1), vbExclamation, DEF_PROJECT_NAME
    
        End If
    
    End If

    '..do small parts first
    If clsNest.CutSmallFirst Then Call m_SmallFirst

    ' 04/20/06 - rg
    '
    ' do onion skinning
    If clsNest.OnionSkin Then Call g_DoOnionSkin

    With Material
      ActiveDrawing.Attribute(DEF_ATT_MATERIAL_X) = clsNest.SheetWidth
      ActiveDrawing.Attribute(DEF_ATT_MATERIAL_Y) = clsNest.SheetLength
      
      If .ProcessWaste Then
        ProcessWasteMaterial .ProcessWasteMCStyle, .ProcessWasteDepthOfCut, .ProcessWasteFinalSheetScrap, _
          .ProcessWasteCutTowardsComponents, .ProcessWasteStrategy, .HorizontalSpacing, .VerticalSpacing
      End If
    End With

    ' Apply any tool sorting
    m_ToolSorting

    ' Remove all nestlists
    N.DeleteAllNestLists

    '..insert report data
    ' Test to see if twin head nesting has been used for this nest
    ' This flag is set to "0" prior to the nest
    ' TwinHeadNesting will set this flag to "1" the nest has been re-optimised
    TwinHeadNestingFlag = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_TWINHEAD, "")

    ' Only produce reports for non-twinhead nested jobs
    ' (Twin head nesting produces its own reports)
    If TwinHeadNestingFlag <> "1" Then
      If Not clsOptions.DisableReports Then
          Call m_InsertReportDataRouter(Material)
      End If
    End If

    ' Test to see if the reverse side nesting add-in is installed and enabled

    ReverseSideNestingFlag = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_REVERSENEST, "")

    ' Set an attribute on the active drawing to indicate that reverse nesting can begin
    ' (this will examined by the reverse side nesting macro)
    ActiveDrawing.Attribute(DEF_ATT_CAN_START_REVERSE_NEST) = "1"

    ' Test to see if the reverse side nesting registry flag
    If ReverseSideNestingFlag = "1" Then

        ' Clear the reverse side nesting registry flag
        SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_REVERSENEST, "0"

        ' Ensure Update rapids is switched off
        SuppressUpdateRapids False

        ' Run the reverse side nesting add-in
        App.Run "AlphaDOOR_ReverseSideNesting.Main.ReverseSideNestingMain"

    End If

    ' Clear the reverse side nesting attribute
    ActiveDrawing.Attribute(DEF_ATT_CAN_START_REVERSE_NEST) = "0"

    ' Ensure Update rapids is switched off
    SuppressUpdateRapids False

    '..save the program
    With clsOptions

        '..set the post variables
        Call m_SetPostVariables(clsNest.SheetWidth, clsNest.SheetLength, _
                                clsNest.SheetThickness, _
                                ((lngNestNumber + 1) * 1000) + 1, .PromptForProgram)

        strNestANCName = mstr_CompileNestOrNCFilename(Material)

        ActiveDrawing.Attribute(DEF_ATT_ALPHADOOR) = "1"
        ActiveDrawing.Attribute(DEF_ATT_NC_FILE_EXTENSION) = .NCFileExtension

        If .OutputToSecondLocation Then
            
            If .OutputSecondNestNCFiles Then
            
                If .OutputSecondCreateSubFolder Then
                  strNestANCFullName = gstr_EnsureBackslash(.PathToRootSecond) & gstr_JobName & "\" & strNestANCName
                Else
                  strNestANCFullName = gstr_EnsureBackslash(.PathToRootSecond) & strNestANCName
                End If
            
                ' Test for MPR output
                If clsOptions.OutputMPR Then
                    With clsNest
                        m_MprSave CStr(.SheetLength), CStr(.SheetWidth), CStr(clsNest.SheetThickness), "0", "0", "0", "0", "0", "0", "0", "", "0", strNestANCFullName & DEF_EXTENSION_MPR
                    End With
                ElseIf clsOptions.OutputNCHops Then
                    With clsNest
                        'm_NCHopsOutputTemp
                        m_NCHopsOutput CStr(.SheetLength), CStr(.SheetWidth), CStr(clsNest.SheetThickness), Material
                    End With
                Else
                    '..output nc to file
                    ActiveDrawing.OutputNC strNestANCFullName & "." & clsOptions.NCFileExtension, acamOutNcFILE, False
                End If
                
                '..collect the nesting information
                Set nstInfo = ActiveDrawing.GetNestInformation
                
                If .SplitNestedPrograms Then
                    
                    Call m_SplitNestANC(Material, nstInfo, strNestANCFullName & "." & clsOptions.NCFileExtension, strNestANCName, colNestANC, True)
        
                End If
            
            End If
        
        End If
        
        If clsOptions.OutputResultsSubFolder Then
          strNestANCFullName = gstr_EnsureBackslash(.PathToRoot) & gstr_JobName & "\" & strNestANCName
        Else
          strNestANCFullName = gstr_EnsureBackslash(.PathToRoot) & strNestANCName
        End If
        
        ' Test for MPR output
        If clsOptions.OutputMPR Then
          With clsNest
            m_MprSave CStr(.SheetLength), CStr(.SheetWidth), CStr(clsNest.SheetThickness), "0", "0", "0", "0", "0", "0", "0", "", "0", strNestANCFullName & DEF_EXTENSION_MPR
          End With
        ElseIf clsOptions.OutputNCHops Then
          With clsNest
            'm_NCHopsOutputTemp
            m_NCHopsOutput CStr(.SheetLength), CStr(.SheetWidth), CStr(clsNest.SheetThickness), Material
          End With
        Else
          '..output nc to file
          ActiveDrawing.OutputNC strNestANCFullName & "." & clsOptions.NCFileExtension, acamOutNcFILE, False
        End If

        '..up the part count (nested files)
        lngNestNumber = lngNestNumber + 1

        '..collect the nesting information
        Set nstInfo = ActiveDrawing.GetNestInformation

        If .SplitNestedPrograms Then
            
            Call m_SplitNestANC(Material, nstInfo, strNestANCFullName & "." & clsOptions.NCFileExtension, strNestANCName, colNestANC, False)

        Else

            If Not .OutputMPR And Not .OutputNCHops Then
              '..added nested anc file to array
              colNestANC.add strNestANCFullName & "." & clsOptions.NCFileExtension
            End If

            '..update the nest anc path in the database
            If Not .DisableReports Then
                Call m_UpdateDBNestPaths(DEF_RST_RPT_NESTANC, DEF_RST_RPT_PATHTONESTANC, strNestANCFullName & "." & clsOptions.NCFileExtension, strNestANCName & "." & clsOptions.NCFileExtension, False)
            End If

        End If

        ' 16 Aug 13
        ' Set the NC Filenames as attributes on the nested sheet geometries (For Alphacam Report generation)
        Set Ni = ActiveDrawing.GetNestInformation
        
        For Each Nsh In Ni.Sheets
          Nsh.Path.Attribute(DEF_ATT_ANC_NAME) = strNestANCName & "." & clsOptions.NCFileExtension
          Nsh.Path.Attribute(DEF_ATT_ANC_FULLNAME) = strNestANCFullName & "." & clsOptions.NCFileExtension
        Next
        
        '..assign the nested ard file name
        strNestARD = mstr_CompileNestOrNCFilename(Material) & DEF_EXTENSION_ARD
        
        '..save to second location?  this is done first so that default path is used for reports    '..07.21.02 - rg
        If .OutputToSecondLocation Then
                
            If clsOptions.OutputSecondCreateSubFolder Then
              strSave = gstr_EnsureBackslash(.PathToRootSecond) & gstr_JobName & "\" & strNestARD
            Else
              strSave = gstr_EnsureBackslash(.PathToRootSecond) & strNestARD
            End If

            If clsOptions.OutputSecondNestedDrawings Then
              ActiveDrawing.SaveAs strSave
            End If
        
        End If
        
        
        '..are we saving it
        If .SaveAllNestARD Then

            '..save the nested drawing
            If clsOptions.OutputResultsSubFolder Then
              ActiveDrawing.SaveAs gstr_EnsureBackslash(.PathToRoot) & gstr_JobName & "\" & strNestARD
            Else
              ActiveDrawing.SaveAs gstr_EnsureBackslash(.PathToRoot) & strNestARD
            End If
            ' 11/21/05 - rg
            '
            Call ActiveDrawing.AddToRecentFileList

            '..update the nest ard path in the database
            If Not .DisableReports Then
                Call m_UpdateDBNestPaths(DEF_RST_RPT_NESTARD, DEF_RST_RPT_PATHTONESTARD, ActiveDrawing.FullName, ActiveDrawing.Name & DEF_EXTENSION_ARD, False)
            End If

        End If

        ' Test to see if the file naming convention is using a sequential order number
        ' rather than the order name
        If Not .UseOrderName Then
          ' increment the sequential number
          If Not gbln_IncrementNCSeqNum Then
            MsgBox Frame.ReadTextFile(strCTX, 500, 124), vbExclamation, DEF_PROJECT_NAME
          End If
        End If


    End With




Controlled_Exit:

End Sub

Public Function mbln_CompareRouterNestToOrder(Router As CRouter, MaterialName As String) As Boolean
  Dim Ni                As NestInformation
  Dim Sheet             As NestSheet
  Dim lQuantity         As Long
  Dim NestPartCount     As Long
  Dim sTwinHeadSheet    As String
  Dim Door              As CDoor
  Dim intNumHeads       As Integer
  Dim strNumHeads       As String
'
  On Error GoTo mbln_CompareRouterNestToOrder_Error
  lQuantity = 0
  
  For Each Door In Router.colMaterials(MaterialName).colDoors
    ' TFS #49412 - Added check for ByPass nesting
    If Not Door.ByPassNest Then
      lQuantity = lQuantity + Door.Quantity
    End If
  Next
    
  Set Ni = ActiveDrawing.GetNestInformation
  ActiveDrawing.Attribute(DEF_ATT_NUM_SHEETS) = Ni.Sheets.Count
    
  For Each Sheet In Ni.Sheets
    strNumHeads = Sheet.Path.Attribute(DEF_ATT_NUM_HEADS)
    If strNumHeads <> "" Then
      intNumHeads = CInt(strNumHeads)
    Else
      intNumHeads = 0
    End If
    If intNumHeads > 0 Then
      NestPartCount = NestPartCount + Sheet.Parts.Count * intNumHeads
    Else
      sTwinHeadSheet = Sheet.Parts(1).Paths(1).Attribute(DEF_ATT_TWIN_HEAD_PATH)
      If sTwinHeadSheet = "1" Then
        NestPartCount = NestPartCount + (Sheet.Parts.Count * 2)
      Else
        NestPartCount = NestPartCount + Sheet.Parts.Count
      End If
    End If
  Next
    
  If NestPartCount = lQuantity Then
    mbln_CompareRouterNestToOrder = True
  End If

Exit Function

mbln_CompareRouterNestToOrder_Error:
  MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
End Function



Private Sub m_MprSave(MprLength As String, MprWidth As String, MprThickness As String, MprProPartX As String, _
                MprProPartY As String, MprUnProPartX As String, MprUnProPartY As String, MprUnProPartZ As String, _
                MprScrapX As String, MprScrapY As String, MprComment As String, MprOnTop As String, mprFileName As String)


    Dim MyVBA As Object 'VBE '
    Dim Project As Object 'VBProject '
    Dim strWWProject As String
    Dim blnFoundWoodWop As Boolean
    
    

    Set MyVBA = App.VBE
    For Each Project In MyVBA.VBProjects
        If InStr(1, Project.Name, "WoodWop4xV20_edge", vbTextCompare) = 1 Then
            strWWProject = "WoodWop4xV20_edge.Make.MprFromExternal"
            blnFoundWoodWop = True
            Exit For
        ElseIf InStr(1, Project.Name, "WoodWop4xV20_normal", vbTextCompare) = 1 Then
            strWWProject = "WoodWop4xV20_normal.Make.MprFromExternal"
            blnFoundWoodWop = True
            Exit For
        End If
    Next

    MprScrapX = clsOptions.MPR_ScrapX
    MprScrapY = clsOptions.MPR_ScrapY
    
    MprProPartX = clsOptions.MPR_ProcPartX
    MprProPartY = clsOptions.MPR_ProcPartY
    
    MprUnProPartX = clsOptions.MPR_UnProcPartX
    MprUnProPartY = clsOptions.MPR_UnProcPartY
    MprUnProPartZ = clsOptions.MPR_UnProcPartZ
    
    If blnFoundWoodWop Then
      ' run sub from Woodwop macro
      App.Run strWWProject, _
                          MprLength & ";" & _
                          MprWidth & ";" & _
                          MprThickness & ";" & _
                          MprProPartX & ";" & _
                          MprProPartY & ";" & _
                          MprUnProPartX & ";" & _
                          MprUnProPartY & ";" & _
                          MprUnProPartZ & ";" & _
                          MprScrapX & ";" & _
                          MprScrapY & ";" & _
                          MprComment & ";" & _
                          MprOnTop & ";" & _
                          mprFileName

    End If

End Sub



Public Sub g_Make_Master(strOrder As String)
    
    Dim FSO                     As Scripting.FileSystemObject
    Dim objAS                   As CAppSettings
    Dim rstOrderProcessed       As ADODB.Recordset
    Dim DrwMaterial             As Material
    Dim colNestANC              As Collection
    Dim colDOORANC              As Collection
    Dim blnNestingExists        As Boolean
    Dim strCurrentPost          As String
    Dim strCurrentAPC           As String
    Dim strPromptMaterial       As String
    Dim strPromptPart           As String
    Dim strJob                  As String
    Dim strMsg                  As String
    Dim strPost                 As String
    Dim lngPartNumber           As Long
    Dim lngOrder                As Long
    Dim strOrders()             As String
    Dim intCount                As Integer
    Dim bNestListStarted        As Boolean
    Dim blnOnion                As Boolean
    Dim Router                  As CRouter
    Dim Material                As CMaterial
    Dim Door                    As CDoor
    Dim lngAlphacimIncID        As Long
    
On Error GoTo g_Make_Master_Error
                
    '..let's assume failure =)
    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_COMPLETENEST, 0
        
    '..get the current post so we can change back later
    strCurrentPost = mstr_CurrentPost
    
    ' Ensure that 3D views are turned off
    ActiveDrawing.ThreeDViews = False
    
    ' 11/21/05 - rg
    '
    ' lock acam
    Call g_LockAcam
    
    
    'Work around for Acam 2011. Geometry group numbers no longer preserved
    'Toolpath are applied to incorrect groups
    'See Salesforce case: https://emea.salesforce.com/5002000000DYiqB
    'Team server work item 41956
    If App.DisableUndo = True Then
      App.DisableUndo = False
    End If
        
    '..set new class objects
    Set clsOptions = New COptions
    Set clsPathData = New CPathData
    Set clsTypeData = New CTypeData
    
    Set FSO = New Scripting.FileSystemObject
    
    Set colNestANC = New Collection
    Set colDOORANC = New Collection
    Set colDeleteFiles = New Collection
    
    strCTX = clsOptions.CTXFile
        
    If colVBAUserStyles.Count = 0 Then
      If Not g_GetVBAProjects Then GoTo Controlled_Exit
    End If
        
    '..make sure the window is visible
    App.Visible = True
        
    '..initialize
    lngPartNumber = 0
    glng_MaterialIndex = 1
    blnOnion = False
    Set colPressReportData = Nothing
    
    gstr_JobIDs = strOrder
    
    If InStr(1, gstr_JobIDs, ",") > 0 Then
      gstr_JobName = "MergedOrder"
      glng_CustomerID = 1
    Else
      gstr_JobName = mstr_GetJobName(CLng(strOrder))
      glng_CustomerID = glng_GetCustomerIDFromOrderID(CLng(strOrder))
    End If
    
    '..let the user know it's starting
    With Frame
        .ShowProgressBox .ReadTextFile(strCTX, 300, 3), .ReadTextFile(strCTX, 300, 19)
    End With
        
    '..make sure metafile directory exists
    If Not FSO.FolderExists(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) Then
        FSO.CreateFolder gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE
    End If
                
    '..Test for output to order specific sub folder
    If clsOptions.OutputResultsSubFolder Then
      If Not FSO.FolderExists(gstr_EnsureBackslash(clsOptions.PathToRoot) & gstr_JobName) Then
        FSO.CreateFolder gstr_EnsureBackslash(clsOptions.PathToRoot) & gstr_JobName
      End If
    End If
    
    If clsOptions.OutputToSecondLocation Then
      If clsOptions.OutputSecondCreateSubFolder Then
        If Not FSO.FolderExists(gstr_EnsureBackslash(clsOptions.PathToRootSecond) & gstr_JobName) Then
          FSO.CreateFolder gstr_EnsureBackslash(clsOptions.PathToRootSecond) & gstr_JobName
        End If
      End If
    End If
    
    '..connect to the database
    If Not gbln_ConnectToDB Then GoTo Controlled_Exit
    
    ' Populate the nesting Zone collection
    m_PopulateNestingZones
    
    ' Populate the Press data for each of the orders being processed
    m_PopulatePressData strOrder
    
    ' Populate the Router data for each of the orders being processed
    m_PopulateRouterData strOrder
    
    strOrders = Split(strOrder, ",")
    
    gbln_UpdateAlphacim = gbln_TestAlphacim
    
    ' Clear out any existing report data from the database
    ' loop through each order
    For intCount = LBound(strOrders) To UBound(strOrders)
        
        lngOrder = strOrders(intCount)
        
        '..wipe out any graphics and database info tied to the current job
        Call m_DeleteReportData(lngOrder, strJob)
    Next
        
    If colPressData.Count > 0 Then
            
      mbln_PopulateReportDataPress strOrder
            
      gbln_Make_Master_Press
    
      '..populate the RouterReportData object
      mbln_PopulateReportData strOrder
    
    End If
    
    
    For Each Router In colRouter
    
      ' Select the post
      ' Test for OnePost
      If UCase(Right(Router.PostProcessor, 7)) <> "ONEPOST" Then
          
          With clsOptions
              
              '..if one has been set then select, else leave alone
              If (Len(.DefaultPost) <> 0) Then
                              
                  '..verify post existance
                  If FSO.FileExists(.DefaultPost) Then
                                         
                          ' 04/20/06 - rg
                          '
                          ' if VBA post then snag the current APC file for the selected
                          ' post this and then reset it just in we need to store case
                          ' what's used for AD is different than what is used by acam
                          '
                          ' note that this must be done prior to selecting the post
                          If .IsPostVBA Then
                                  
                                  Set objAS = New CAppSettings
                                  
                                  strPost = gs_ParseFileName(.DefaultPost, False)
                                          
                                  ' check and load apc file
                                  '
                                  ' this gets the default apc file set within alphacam - may or may not exist
                                  ' we store it so that we can reset it later
                                  objAS.AppName = "PSPostExt"
                                  strCurrentAPC = objAS.ReadEntry("APCSettings\" & strPost, "LastAPC", vbNullString)
                          
                                  ' now save APC file to be used by AD
                                  Call objAS.WriteEntry("APCSettings\" & strPost, "LastAPC", .APCFile)
                          
                          End If
                          
                          ' 04/20/06 - rg
                          '
                          ' only select post if not already active
                          If (StrComp(App.PostFileName, .DefaultPost, vbTextCompare) <> 0) Then
                                  Call App.SelectPost(.DefaultPost)
                                  DoEvents
                          End If
                      
                  Else
                      
                      '..setup the prompt
                      strMsg = Frame.ReadTextFile(strCTX, 500, 83) & Space(3) & vbCrLf & _
                               Frame.ReadTextFile(strCTX, 500, 84) & Space(3)
                      
                      '..let them know that the current alphacam post will be used
                      MsgBox strMsg, vbInformation, DEF_PROJECT_NAME
                          
                  End If
                  
              End If
          
          End With
                  
      End If
    
      ' Get the rules for this router
      m_PopulateRulesCollection Router
      
      ' Loop through the materials for this router
      For Each Material In Router.colMaterials
      
          '..need new drawing
          App.New
          
          bNestListStarted = False
    
          '..set the material
          Set DrwMaterial = ActiveDrawing.GetMaterial
          DrwMaterial.Name = Material.MaterialName
          ActiveDrawing.SetMaterial DrwMaterial
    
          ' Set a global attribute to indicate the current material being nested
          ActiveDrawing.Attribute(DEF_ATT_MATERIAL_NAME) = Material.MaterialName
          
          '..set the progress text
          strPromptMaterial = Frame.ReadTextFile(strCTX, 300, 5) & Space(1) & _
                              Material.MaterialName
    
          Frame.SetProgressText strPromptMaterial
          DoEvents
                
          Set clsNest = New CNest
          
          clsNest.StartNestListRouter gstr_JobName, Material, False, False
          
          ' TFS #49411
          '..initialize nesting flag
          blnNestingExists = False
          
          For Each Door In Material.colDoors
              
              If clsNest.NestingOption <> adoorNESTING_DISABLED Then
                '..create the nest list header
                If Not bNestListStarted Then
                    bNestListStarted = True
                End If
              Else
                ' Set this door to bypass nesting if the global setting is for nesting disabled
                Door.ByPassNest = True
              End If
              
              ' TFS #49411 - Moved to outside of loop
              '..initialize nesting flag
              'blnNestingExists = False
    
              '..set the current type
              clsTypeData.TypeName = Door.TypeName
              clsTypeData.TypeStyle = Door.StyleNumber

              '..set the progress text
              strPromptPart = Frame.ReadTextFile(strCTX, 300, 6) & Space(1) & clsTypeData.TypeName & vbCrLf
    
              With Frame
                  .ShowProgressBox .ReadTextFile(strCTX, 300, 3), strPromptPart & strPromptMaterial
              End With
    
              DoEvents

              '..make the part
              If Not mbln_ProcessPart(Door, lngPartNumber, colDOORANC, False, False) Then

                  '..had an issue
                  SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_COMPLETENEST, 0

                  '..let the user know that something wasn't right
                  MsgBox Frame.ReadTextFile(strCTX, 600, 26) & Space(3), vbInformation, DEF_PROJECT_NAME

                  App.New
                  GoTo Controlled_Exit

              Else
              
                  If gbln_UpdateAlphacim Then
                                          
                      lngAlphacimIncID = mlng_GetCIM_INCID_Value(Door)
                    
                      If gbln_AlphacimAssemblyItemExists(lngAlphacimIncID) Then
                      
                          If Not gbln_UpdateAlphacimComponentGrouping(lngAlphacimIncID, gstr_ConvertToAlphaGrouping(Door.ComponentGrouping)) Then
                              
                              MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 142) & Chr(13) & _
                                "Alphacim INC_ID=" & lngAlphacimIncID & Chr(13) & _
                                "Component Grouping=" & Door.ComponentGrouping, vbExclamation, DEF_PROJECT_NAME
                              
                          End If
                      
                      End If
                    
                  End If
              
              End If
    
              '..if nesting then set flag
              If Not Door.ByPassNest Then
                  blnNestingExists = True
              End If
          
          
          ' Get the next door
          Next
      
          If bNestListStarted Then
  
              '..start clean so we can nest
              App.New
  
              '..if no parts have been created then bail
              If CBool(lngPartNumber) Then
                  
                  '..are we nesting?
                  If blnNestingExists Then
                  
                      m_DoNestingRouter Router, Material, colNestANC
                  
                  End If
              End If
          
          End If
      
      
      ' Get the Next material
      Next
    
    ' Get the next Router
    Next
    
    Dim PressReportData         As CPressReportData
    Dim PressReportSheet        As CPressReportSheet
    Dim PressReportDataItem     As CPressReportDataItem
    Dim PressReportQty          As CPressReportQty
    Dim lngCount                As Long
    Dim colReportPK             As New Collection
    Dim lngPKCount              As Long
    Dim strKey                  As String

    For Each PressReportData In colPressReportData
      For Each PressReportSheet In PressReportData.colSheets
        For Each PressReportDataItem In PressReportSheet.colComponents
          
          strKey = PressReportData.DetailID & "_" & PressReportSheet.SheetName & "_" & PressReportDataItem.RouterSheetName
          
          If Not gbln_ItemExistsInCollection(colReportPK, strKey) Then
            Set PressReportQty = New CPressReportQty
            With PressReportQty
              .Quantity = 1
              .DetailID = PressReportData.DetailID
              .RouterSheet = PressReportDataItem.RouterSheetName
              .colReportID.add PressReportDataItem.ReportID
            End With
            colReportPK.add PressReportQty, strKey
          Else
            With colReportPK(strKey)
              .Quantity = .Quantity + 1
              .colReportID.add PressReportDataItem.ReportID
            End With
          End If
        Next
      Next
    Next

    For Each PressReportQty In colReportPK
      For lngPKCount = 1 To PressReportQty.colReportID.Count
        gdb_CDM.Execute "UPDATE AD_REPORT_DATA SET PressQuantityThisSheet=" & PressReportQty.Quantity & " WHERE PK=" & PressReportQty.colReportID(lngPKCount)
        'Debug.Print PressReportQty.colReportID(lngPKCount), PressReportQty.Quantity
      Next
    Next
        
    '..were there any parts processed?
    If CBool(lngPartNumber) Then

        '..output to alphaedit?
        If Not clsOptions.OutputToNCOnly Then Call m_OutputToAlphaEDIT(colNestANC, colDOORANC)

        '..looks like we made it so let the dll know
        SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_COMPLETENEST, -1

        '..force focus to alphacam
        App.Visible = True

        With Frame

            '..get rid of the little box
            .CloseProgressBox
            
            Call g_UnlockAcam(True)

            DoEvents

            '..let the user know that the nest is complete and ask to shut down alphadoor
            If MsgBox(.ReadTextFile(strCTX, 600, 24) & Space(1) & vbCrLf & vbCrLf & _
                      .ReadTextFile(strCTX, 600, 25) & Space(3), vbQuestion + vbYesNo, _
                      .ReadTextFile(strCTX, 120, 1)) = vbNo Then

                '..tell the dll we want to create another
                SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_CREATEANOTHER, -1

            Else

                '..tell the dll we don't want to create another
                SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_CREATEANOTHER, 0

            End If

        End With

        '..update processed date field for each order                                             '..07.08.02 - rg
        For intCount = LBound(strOrders) To UBound(strOrders)

            lngOrder = strOrders(intCount)
            Set rstOrderProcessed = grst_GetOrderInfo(lngOrder)

            With rstOrderProcessed

                If Not rstOrderProcessed Is Nothing Then
                    .Fields!ProcessedDate = CStr(Now)
                    .Update
                End If
                If (.State = adStateOpen) Then .Close

            End With
        Next

    Else

        '..looks like we made it so let the dll know
        SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_COMPLETENEST, 0

        '..force focus to alphacam
        App.Visible = True

        '..get rid of the little box
        Frame.CloseProgressBox

        Call g_UnlockAcam(True)

        DoEvents

        '..let the user know that nothing was done
        MsgBox Frame.ReadTextFile(strCTX, 600, 138) & Space(3), vbInformation, DEF_PROJECT_NAME

    End If

    DoEvents

Controlled_Exit:

On Error Resume Next

    '..get rid of the little box
    Frame.CloseProgressBox

    If (StrComp(App.PostFileName, "ONEPOST", vbTextCompare) <> 0) Then

            ' 04/20/06 - rg
            '
            ' restore the original APC file to be used by acam
            '
            ' note that this must be done prior to reslecting the original post
            If clsOptions.IsPostVBA Then
                    If (objAS Is Nothing) Then Set objAS = New CAppSettings
                    Call objAS.WriteEntry("APCSettings\" & strPost, "LastAPC", strCurrentAPC)
            End If

            ' only reselect post if changed from original
            If (StrComp(App.PostFileName, strCurrentPost, vbTextCompare) <> 0) Then
                    If FSO.FileExists(strCurrentPost) Then Call App.SelectPost(strCurrentPost)
            End If

    End If
        
    ' TFS #56788 - Clear current drawing if required
    ' Do this before any file clean-up
    If Not clsOptions.KeepLastDrawingOpen Then
      App.New
    End If
    
    ' Remove any unwanted files
    On Error Resume Next
    For lngCount = 1 To colDeleteFiles.Count
      FSO.DeleteFile colDeleteFiles(lngCount)
    Next
    On Error GoTo 0

    Set clsNest = Nothing
    Set clsOptions = Nothing
    Set clsTypeData = Nothing
    Set clsPathData = Nothing
    Set FSO = Nothing
    Set colDOORANC = Nothing
    Set colNestANC = Nothing
    Set colDeleteFiles = Nothing
        
    '..unload all the forms
    Call g_UnLoadAllForms
        
    ' 11/21/05 - rg
    '
    ' start new drawing and refresh
    'Call App.New
    Call g_UnlockAcam(True)
    
    If (Err.Number <> 0) Then Err.Clear
    
Exit Sub

g_Make_Master_Error:
    
    '..had an issue
    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_COMPLETENEST, 0
    
    MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
    MsgBox Frame.ReadTextFile(strCTX, 600, 26) & Space(3), vbInformation, DEF_PROJECT_NAME
    If (Err.Number <> 0) Then WriteError Err, True, "g_Make_Master"
    Resume Controlled_Exit

End Sub

Public Sub m_PopulatePressData(OrderString As String)
  Dim strOrders()       As String
  Dim strOrder          As String
  Dim intOrderCount     As Integer
  Dim rstOrderDetails   As ADODB.Recordset
  Dim rstOrderInfo      As ADODB.Recordset
  Dim Press             As CPress
  Dim PressColour       As CPressColour
  Dim PressThickness    As CPressThickness
  Dim PressDoor         As CDoor
  Dim lngPressID        As Long
  Dim lngColourID       As Long
  Dim lngOrder          As Long
  Dim strJob            As String
  

  Set colPressData = Nothing
  
  strOrders = Split(OrderString, ",")
  
  For intOrderCount = LBound(strOrders) To UBound(strOrders)
    strOrder = strOrders(intOrderCount)
    
    Set rstOrderDetails = grst_GetOrderDetailsSpecific(CLng(strOrder))
        
    If Not rstOrderDetails Is Nothing Then
      
      Do While Not rstOrderDetails.EOF
        
        If gbln_SQLServer Then
          lngPressID = gvar_CheckNull(rstOrderDetails.Fields!PressID)
        Else
          lngPressID = gvar_CheckNull(rstOrderDetails.Fields![AD_ORDER_DETAILS.PressID])
        End If
        If lngPressID > 0 Then
          If gbln_ItemExistsInCollection(colPressData, rstOrderDetails.Fields!PressName) Then
            Set Press = colPressData(rstOrderDetails.Fields!PressName)
          Else
            Set Press = New CPress
            With Press
              .PressID = lngPressID
              .PressName = gvar_CheckNull(rstOrderDetails.Fields!PressName)
              .PressLength = gvar_CheckNull(rstOrderDetails.Fields!PressLength)
              .PressWidth = gvar_CheckNull(rstOrderDetails.Fields!PressWidth)
              .GapAtEdge = gvar_CheckNull(rstOrderDetails.Fields!GapAtSheetEdge)
              .GapBetweenParts = gvar_CheckNull(rstOrderDetails.Fields!MinGapBetweenPaths)
              .NumberBySize = gvar_CheckNull(rstOrderDetails.Fields!NumberComponentsBySize)
              .PartRotation = gvar_CheckNull(rstOrderDetails.Fields!PartRotation)
              .UseTrueShape = gvar_CheckNull(rstOrderDetails.Fields!UseTrueShape)
              .TrueShapePackTo = gvar_CheckNull(rstOrderDetails.Fields!TrueShapePackTo)
              If IsNull(rstOrderDetails.Fields!RectPackTo) Then
                .RectPackTo = 2
              Else
                .RectPackTo = gvar_CheckNull(rstOrderDetails.Fields!RectPackTo)
              End If
              colPressData.add Press, .PressName
            End With
          End If
        
          lngColourID = gvar_CheckNull(rstOrderDetails.Fields!ColourID)
          If lngColourID > 0 Then
            If gbln_ItemExistsInCollection(Press.colPressColours, rstOrderDetails.Fields!ColourName) Then
              Set PressColour = Press.colPressColours(rstOrderDetails.Fields!ColourName)
            Else
              Set PressColour = New CPressColour
              With PressColour
                .ColourID = lngColourID
                .ColourName = rstOrderDetails.Fields!ColourName
                Press.colPressColours.add PressColour, PressColour.ColourName
              End With
            End If
              
            If gbln_ItemExistsInCollection(PressColour.colPressThicknesses, "k" & rstOrderDetails.Fields!Thickness) Then
              Set PressThickness = PressColour.colPressThicknesses("k" & rstOrderDetails.Fields!Thickness)
            Else
              Set PressThickness = New CPressThickness
              With PressThickness
                .PressThickness = rstOrderDetails.Fields!Thickness
                PressColour.colPressThicknesses.add PressThickness, "k" & .PressThickness
              End With
            End If
              
            Set PressDoor = New CDoor
            PressDoor.PressName = Press.PressName
            PressDoor.FoilColour = PressColour.ColourName
            PressDoor.DoorThickness = PressThickness.PressThickness
            PressDoor.RotationMethodPress = Press.PartRotation
            
            m_FillDoorData PressDoor, rstOrderDetails, False
            
            lngOrder = CLng(strOrders(intOrderCount))
            
            ' get the job name
            strJob = mstr_GetJobName(lngOrder)
            
            '..get the order info
            Set rstOrderInfo = grst_GetOrderInfo(lngOrder)
            
            '..fill up the customer and door type information information
            Call m_FillOrderInfo(PressDoor, rstOrderInfo, strJob, lngOrder)
            
            PressThickness.colPressComponents.add PressDoor, "k" & PressDoor.DetailID
              
          End If
        End If
        rstOrderDetails.MoveNext
      Loop
    End If
  Next

End Sub



Public Sub m_PopulateRouterData(OrderString As String)
  Dim strOrders()       As String
  Dim strOrder          As String
  Dim intOrderCount     As Integer
  Dim rstOrderDetails   As ADODB.Recordset
  Dim rstMaterial       As ADODB.Recordset
  Dim rstOrderInfo      As ADODB.Recordset
  Dim Router            As CRouter
  Dim Material          As CMaterial
  Dim RouterDoor        As CDoor
  Dim strMaterial       As String
  Dim strPostProcessor  As String
  Dim lngOrder          As Long

  Set colRouter = Nothing
  
  strOrders = Split(OrderString, ",")
  
  For intOrderCount = LBound(strOrders) To UBound(strOrders)
    strOrder = strOrders(intOrderCount)
    
    Set rstOrderDetails = grst_GetOrderDetailsSpecific(CLng(strOrder))
        
    If Not rstOrderDetails Is Nothing Then
      
      Do While Not rstOrderDetails.EOF
        
        strPostProcessor = gvar_CheckNull(rstOrderDetails.Fields!PostProcessor)
        
        If strPostProcessor = "" Then
          ' No post specified - use default
          strPostProcessor = gstr_StripLicomDatPath(clsOptions.DefaultPost)
        End If
        
        ' Get the router
        If gbln_ItemExistsInCollection(colRouter, strPostProcessor) Then
          Set Router = colRouter(strPostProcessor)
        Else
          Set Router = New CRouter
          With Router
            .PostProcessor = App.LicomdatPath & strPostProcessor
            colRouter.add Router, strPostProcessor
          End With
        End If
        
        ' Get the material
        strMaterial = gvar_CheckNull(rstOrderDetails.Fields!Material)
        
        If gbln_ItemExistsInCollection(Router.colMaterials, strMaterial) Then
          Set Material = Router.colMaterials(strMaterial)
        Else
          Set rstMaterial = grst_GetMaterial(strMaterial)
          If Not rstMaterial Is Nothing Then
            Set Material = New CMaterial
            m_FillMaterialData Material, rstMaterial
            Router.colMaterials.add Material, strMaterial
          End If
        End If

        lngOrder = CLng(strOrders(intOrderCount))
        
        '..get the order info
        Set rstOrderInfo = grst_GetOrderInfo(lngOrder)
        
        ' Populate the door data
        Set RouterDoor = New CDoor
        m_FillOrderInfo RouterDoor, rstOrderInfo, gstr_JobName, lngOrder
        m_FillDoorData RouterDoor, rstOrderDetails, False
        Material.colDoors.add RouterDoor, "k" & RouterDoor.DetailID
        
        rstOrderDetails.MoveNext
      Loop
      
    End If
  Next

End Sub


Public Function mbln_PopulateReportData(OrderString As String) As Boolean
  Dim strOrders()             As String
  Dim strOrder                As String
  Dim intOrderCount           As Integer
  Dim rstReportDetails        As ADODB.Recordset
  Dim lngDetailID             As Long
  Dim lngReportItemPK         As Long
  Dim RouterReportData        As CRouterReportData
  Dim RouterReportDataItem    As CRouterReportDataItem
  
  Set colReportData = Nothing
  
  strOrders = Split(OrderString, ",")
  
  For intOrderCount = LBound(strOrders) To UBound(strOrders)
    strOrder = strOrders(intOrderCount)
    
    Set rstReportDetails = grst_GetReportDetails(CLng(strOrder))
        
    If rstReportDetails Is Nothing Then
      MsgBox Frame.ReadTextFile(strCTX, 500, 132) & Chr(13) & _
              Frame.ReadTextFile(strCTX, 500, 133), vbExclamation, DEF_PROJECT_NAME
      
      mbln_PopulateReportData = False
      GoTo Controlled_Exit
    End If
    
    rstReportDetails.MoveFirst
    
    Do While Not rstReportDetails.EOF
      lngDetailID = rstReportDetails.Fields!DetailID
      If Not gbln_ItemExistsInCollection(colReportData, "k" & lngDetailID) Then
        Set RouterReportData = New CRouterReportData
        colReportData.add RouterReportData, "k" & lngDetailID
      Else
        Set RouterReportData = colReportData("k" & lngDetailID)
      End If
    
      lngReportItemPK = rstReportDetails.Fields!PK
      
      Set RouterReportDataItem = New CRouterReportDataItem
      With RouterReportDataItem
        .ReportDataPK = lngReportItemPK
        .DetailID = lngDetailID
      End With
          
      RouterReportData.colRouterReportData.add RouterReportDataItem, "k" & lngReportItemPK
      
      rstReportDetails.MoveNext
    Loop
    
  Next
    
Controlled_Exit:

  If Not rstReportDetails Is Nothing Then
    If rstReportDetails.State = adStateOpen Then
      rstReportDetails.Close
    End If
  End If

  Set rstReportDetails = Nothing
  
Exit Function

mbln_PopulateReportData_Error:

  mbln_PopulateReportData = False
  MsgBox Err.Description, vbExclamation, "mbln_PopulateReportData"
  Resume Controlled_Exit

End Function




Public Function mbln_PopulateReportDataPress(OrderString As String) As Boolean
  Dim strOrders()             As String
  Dim strOrder                As String
  Dim intOrderCount           As Integer
  Dim rstOrderDetails         As ADODB.Recordset
  Dim lngDetailID             As Long
  Dim PressReportData         As CPressReportData
    
  strOrders = Split(OrderString, ",")
  
  For intOrderCount = LBound(strOrders) To UBound(strOrders)
    strOrder = strOrders(intOrderCount)
    
    Set rstOrderDetails = grst_GetOrderDetailsByID(CLng(strOrder))
        
    If rstOrderDetails Is Nothing Then
      MsgBox "Error", vbExclamation, DEF_PROJECT_NAME
      
      mbln_PopulateReportDataPress = False
      GoTo Controlled_Exit
    End If
    
    rstOrderDetails.MoveFirst
    
    Do While Not rstOrderDetails.EOF
      lngDetailID = rstOrderDetails.Fields!PK
      If Not gbln_ItemExistsInCollection(colPressReportData, "k" & lngDetailID) Then
        
        Set PressReportData = New CPressReportData
        PressReportData.DetailID = lngDetailID
        colPressReportData.add PressReportData, "k" & lngDetailID
      
      End If
          
      rstOrderDetails.MoveNext
    Loop
    
  Next
    
Controlled_Exit:

  If rstOrderDetails.State = adStateOpen Then
    rstOrderDetails.Close
  End If

  Set rstOrderDetails = Nothing
  
Exit Function

mbln_PopulateReportData_Error:

  mbln_PopulateReportDataPress = False
  MsgBox Err.Description, vbExclamation, "mbln_PopulateReportDataPress"
  Resume Controlled_Exit

End Function








Private Function mstr_CompileNestOrNCFilename(Material As CMaterial) As String
  Dim strFilename As String
  Dim lngNCSeqNum As Long
'
  If clsOptions.UseNestedPrefix Then
    strFilename = clsOptions.NestPrefix & DEF_UNDERSCORE
  End If
  
  If clsOptions.UseOrderName Then
    strFilename = strFilename & gstr_JobName
  Else
    ' get the sequential number from the database
    lngNCSeqNum = glng_GetNCSeqNum
    strFilename = strFilename & lngNCSeqNum
  End If
  
  If clsOptions.UseMaterialData Then
    strFilename = strFilename & "_"
    If clsOptions.UseMaterialName Then
      strFilename = strFilename & Material.MaterialName
    Else
      strFilename = strFilename & Material.AbsolutePosition
    End If
  End If

  mstr_CompileNestOrNCFilename = strFilename

End Function


Private Function mstr_CompileNestOrNCFilenamePress(PressDetails As String) As String
  Dim strFilename As String
  Dim lngNCSeqNum As Long
'
  If clsOptions.UseNestedPrefix Then
    strFilename = clsOptions.NestPrefix & DEF_UNDERSCORE
  End If
  
  If clsOptions.UseOrderName Then
    strFilename = strFilename & gstr_JobName
  Else
    ' get the sequential number from the database
    lngNCSeqNum = glng_GetNCSeqNum
    strFilename = strFilename & lngNCSeqNum
  End If
  
  strFilename = strFilename & "_" & PressDetails

  mstr_CompileNestOrNCFilenamePress = strFilename

End Function



Private Function mlng_GetCIM_INCID_Value(Door As CDoor) As Long

  On Error Resume Next
  
  Select Case gstr_AlphacimPKField
    Case "CSV_CustomerName"
      mlng_GetCIM_INCID_Value = CLng(Door.CSV_CustomerName)
    Case "CSV_OrderNumber"
      mlng_GetCIM_INCID_Value = CLng(Door.CSV_OrderNumber)
    Case "CSV_ItemNumber"
      mlng_GetCIM_INCID_Value = CLng(Door.CSV_ItemNumber)
    Case "ProductionComment"
      mlng_GetCIM_INCID_Value = CLng(Door.ProductionComment)
    Case "CustomField1"
      mlng_GetCIM_INCID_Value = CLng(Door.CustomField1)
    Case "CustomField2"
      mlng_GetCIM_INCID_Value = CLng(Door.CustomField2)
  End Select

End Function

Private Function mstr_GetCustomMacroName(MacroFilename As String) As String
    
    Dim MyVBA               As Object
    Dim Project             As Object
    Dim strProjectName      As String
    Dim strMacroFilename    As String
    Dim strProjectFilename  As String

    On Error Resume Next

    strMacroFilename = gstr_ParseName(MacroFilename)
    Set MyVBA = App.VBE
    For Each Project In MyVBA.VBProjects
        ' Try to get the Project filename
        strProjectFilename = Project.FileName
        If (Err.Number = 76) Then
          ' Error 76 will be raised for unsaved projects (Project.Filename will be "")
          Err.Clear
        Else
          If InStr(1, Project.FileName, strMacroFilename, vbTextCompare) > 0 Then
              strProjectName = Project.Name
              Exit For
          End If
        End If
    Next
    
    gstr_CustomMacroFileName = MacroFilename
    mstr_GetCustomMacroName = strProjectName

End Function

Public Function mstr_GetJobName(OrderID As Long) As String
    Dim rstJobName              As ADODB.Recordset
'
    Set rstJobName = grst_GetJobName(OrderID)
    
    If Not (rstJobName Is Nothing) Then
        
        mstr_GetJobName = gvar_CheckNull(rstJobName.Fields!JobName)
    
        If (rstJobName.State = adStateOpen) Then rstJobName.Close
        Set rstJobName = Nothing
    
    End If

End Function


Public Sub g_Preview_Master(lType As Long)  ' sType As String)
    
    Dim FSO                     As Scripting.FileSystemObject
    'Dim sCTX                    As String
    Dim rstType                 As ADODB.Recordset
    Dim lngDummy                As Long
    Dim colDummy                As Collection
    Dim pDummy                  As Path
    Dim Door                    As CDoor
    
On Error GoTo g_Preview_Master_Error
                
    '..set new class objects
    Set clsOptions = New COptions
    Set clsPathData = New CPathData
    Set clsTypeData = New CTypeData
    Set clsNest = New CNest
    
    Set FSO = New Scripting.FileSystemObject
    
    'Frame.ProjectBarUpdating = False
    
    strCTX = clsOptions.CTXFile
        
    ' Get the collection of VBA user styles
    If colVBAUserStyles.Count = 0 Then
      If Not g_GetVBAProjects Then GoTo Controlled_Exit
    End If
        
    '..make sure the window is visible
    App.Visible = True
    
    '..let the user know it's starting
    With Frame
        .ShowProgressBox .ReadTextFile(strCTX, 500, 91), .ReadTextFile(strCTX, 500, 92)
    End With
    
    '..connect to the database
    If Not gbln_ConnectToDB Then GoTo Controlled_Exit
                             
    '..get the door type
    Set rstType = grst_GetDoorTypeData(lType)
        
    '..if nothing here then go get the next
    If Not (rstType Is Nothing) Then
            
            ' 28 feb 12 TFS#49189
            '   + MOVED from above and reinstated
            '
            ' 11/21/05 - rg
            '
            ' lock acam
            Call g_LockAcam
    
            Set Door = New CDoor
            m_FillDoorData Door, rstType, True
            
            '..need new drawing
            App.New
                                                        
            With rstType
              
                '..make the part
                If Not mbln_ProcessPart(Door, lngDummy, colDummy, False, True) Then
                
                    '..let the user know that something wasn't right
                    MsgBox Frame.ReadTextFile(strCTX, 500, 93) & Space(3), vbExclamation, DEF_PROJECT_NAME
                        
                    App.New
                    GoTo Controlled_Exit
                    
                End If
        
            End With
    End If

    ActiveDrawing.Redraw
    DoEvents
                
    With Frame
                
        '..get rid of the little box
        .CloseProgressBox
        
        ' 28 feb 12 TFS#49189
        '   + REINSTATED so we can see the darn thing
        '
        ' 01/19/06 - rg
        '
        ' unlock acam
        Call g_UnlockAcam(True)
    
        Set pDummy = ActiveDrawing.UserSelectOneGeo(.ReadTextFile(strCTX, 500, 94))

        Do Until ActiveDrawing.UserAction = acamUserESCAPE
            Set pDummy = ActiveDrawing.UserSelectOneGeo(.ReadTextFile(strCTX, 500, 94))
            DoEvents
        Loop
   
    End With
                
Controlled_Exit:
                
On Error Resume Next
                                    
    Frame.CloseProgressBox
    App.New
                                                
    Set rstType = Nothing
    Set clsOptions = Nothing
    Set clsTypeData = Nothing
    Set clsPathData = Nothing
    Set clsNest = Nothing
    Set FSO = Nothing
    Set pDummy = Nothing
    
    '..unload all the forms
    Call g_UnLoadAllForms
    
    ' 28 feb 12 TFS#49189
    '   + REINSTATED with additional check
    '
    ' 11/21/05 - rg
    '
    ' unlock acam
    If Not App.ActiveDrawing.ScreenUpdating Then Call g_UnlockAcam(True)
    
    If (Err.Number <> 0) Then Err.Clear
    
Exit Sub

g_Preview_Master_Error:
    
    MsgBox Frame.ReadTextFile(strCTX, 600, 26) & Space(3), vbInformation, DEF_PROJECT_NAME
    If (Err.Number <> 0) Then WriteError Err, True, "g_Preview_Master"
    Resume Controlled_Exit

End Sub

Public Sub m_CreateAlphaCAMDrawingsOfSheetsPress(PressDetails As String)

    Dim nInfo                   As NestInformation
    Dim SH                      As NestSheet
    Dim SheetPath               As Path
    Dim DrawingText             As Text
    Dim SheetDrawing            As Drawing
    Dim strNestImage            As String
    Dim strNestARD              As String
    Dim TextPaths               As Paths
    Dim TextPath                As Path
    Dim lBackgroundColour       As Long
    Dim iNumSheets              As Integer
    Dim iCount                  As Integer
    Dim sEMFPath                As String
    Dim colSheetNames           As New Collection
    Dim SheetIdent              As String
    Dim Md                      As MillData
    Dim SheetMinX               As Double
    Dim SheetMinY               As Double
    Dim SheetMaxX               As Double
    Dim SheetMaxY               As Double
'
    Set nInfo = ActiveDrawing.GetNestInformation
    
    iNumSheets = nInfo.Sheets.Count
    
'    With nInfo
'        For Each SH In nInfo.Sheets
'
'            lngPartCount = 1
'
'            For Each Npi In SH.Parts
'
'              strNestImage = gstr_CheckDir(Frame.PathOfThisAddin & DEF_BACKSLASH & DEF_PATH_IMAGE) & _
'                                           gstr_JobName & DEF_UNDERSCORE & PressDetails & DEF_UNDERSCORE & SH.Name & DEF_UNDERSCORE & lngPartCount
'
'              For Each NestPath In Npi.Paths
'                NestPath.Attribute(DEF_ATT_NEST_DOOR_IMAGE) = strNestImage
'                NestPath.Attribute(DEF_ATT_NEST_DOOR_COUNT) = lngPartCount
'              Next
'
'              lngPartCount = lngPartCount + 1
'
'            Next
'
'        Next
'
'    End With
    
    With clsOptions
        strNestARD = mstr_CompileNestOrNCFilenamePress(PressDetails) & DEF_EXTENSION_ARD
        'strNestARD = DEF_NEST_PREFIX & udtCCI.JobID & DEF_UNDERSCORE & clsNest.SheetName & DEF_EXTENSION_ARD
        Frame.ShowProgressBox Frame.ReadTextFile(strCTX, 300, 22), Frame.ReadTextFile(strCTX, 300, 23)
        
        If .OutputResultsSubFolder Then
          ActiveDrawing.SaveAs gstr_EnsureBackslash(.PathToRoot) & gstr_JobName & "\" & strNestARD
        Else
          ActiveDrawing.SaveAs gstr_EnsureBackslash(.PathToRoot) & strNestARD
        End If
    End With
  
    ' Convert all text on the active drawing to geometry
    For Each DrawingText In ActiveDrawing.Text
        SheetIdent = DrawingText.Attribute(attSheetIdent)
        Set TextPaths = DrawingText.ConvertToGeometry
        If SheetIdent <> "" Then
            For Each TextPath In TextPaths
                TextPath.Attribute(attSheetIdent) = SheetIdent
            Next
        End If
    Next
    
    With nInfo
        For Each SH In nInfo.Sheets
            
            Set SheetDrawing = App.CreateTempDrawing
            SheetIdent = gstr_GetNestedSheetIdent(SH.Name)
                        
            Frame.SetProgressText Frame.ReadTextFile(DEF_TEXT, 300, 13) & Space(1) & SH.Name
            
            For Each SheetPath In SH.Paths
                If SheetPath.IsToolPath Then
                    Set Md = SheetPath.GetMillData
                    If Md.ProcessType = 2 Or Md.ProcessType = 4 Or Md.ProcessType = 5 Then
                      SheetPath.Attribute(DEF_ATT_POCKET_PATH) = "1"
                    End If
                End If
                SheetPath.SetLayer ActiveDrawing.Layers(1)
                SheetPath.MoveToDrawing SheetDrawing
            Next
            
            For Each TextPath In ActiveDrawing.Geometries
             
                  
                If TextPath.TestInsidePath(SH.Path) = acamResultTRUE _
                  Or TextPath.Attribute(attSheetIdent) = SheetIdent Then
                    
                    TextPath.SetLayer ActiveDrawing.Layers(1)
                    TextPath.MoveToDrawing SheetDrawing
                
                End If
        
            Next
            
            SH.Path.Sheet = False
            SH.Path.MoveToDrawing SheetDrawing
            
            ' Add the sheet name to a local collection (for language inconsistencies)
            colSheetNames.add SH.Name
                        
'            SheetDrawing.Attribute(DEF_ATT_SHEET_DOOR_COUNT) = SH.Parts.Count
            
            strNestImage = gstr_CheckDir(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) & _
                                         gstr_JobName & DEF_UNDERSCORE & PressDetails & DEF_UNDERSCORE & SH.Name
            
            
            SheetDrawing.SaveAs strNestImage & DEF_EXTENSION_ARD
          
        Next
    End With

    Frame.SetProgressText Frame.ReadTextFile(strCTX, 300, 24)

    
    ' ljo - 15/7/11
    ' Prevents invisible report images with new Alphacam shading modes
    ' TFS#45271
    Dim blnGradientFillWire     As Boolean
    Dim blnGradientFillShaded   As Boolean

    ' Save the current background settings
    blnGradientFillWire = ActiveDrawing.BackgroundColorWireGradient
    blnGradientFillShaded = ActiveDrawing.BackgroundColorShadingGradient
    
    ' Save the current background colour
    ActiveDrawing.BackgroundColorShadingGradient = False
    ActiveDrawing.BackgroundColorWireGradient = False
    lBackgroundColour = ActiveDrawing.BackgroundColor
    ActiveDrawing.BackgroundColor = acamWHITE
'    ActiveDrawing.RedrawShadedViews
    
    For iCount = 1 To iNumSheets
      
      ' Retrieve the AlphaCAM drawing using the name stored in the local collection
      strNestARD = gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE & DEF_BACKSLASH & _
        gstr_JobName & DEF_UNDERSCORE & PressDetails & DEF_UNDERSCORE & colSheetNames(iCount)
    
      App.OpenDrawing strNestARD & DEF_EXTENSION_ARD
    
      For Each SheetPath In ActiveDrawing.Geometries
        If SheetPath.Attribute(DEF_ATT_POCKET_PATH) = "1" Then
          SheetPath.Color = acamYELLOW
        Else
          SheetPath.Color = acamBLACK
        End If
'        SheetPath.Redraw
      Next
      
      sEMFPath = gstr_CheckDir(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) & _
        gstr_JobName & DEF_UNDERSCORE & PressDetails & DEF_UNDERSCORE & colSheetNames(iCount)
            
      ' 04/20/06 - rg
      '     + ADDED so that entire drawing is present in report, this
      '             was not an issue until I added screen locking
      '
      'Call ActiveDrawing.AutoScale  ' .ZoomAll
    
      ' Zoom in based on the sheet extents
      ActiveDrawing.GetExtent SheetMinX, SheetMinY, 0, SheetMaxX, SheetMaxY, 0
      ActiveDrawing.ZoomToBox SheetMinX, SheetMinY, SheetMaxX, SheetMaxY, 2

      ActiveDrawing.SaveEmfFile sEMFPath & DEF_EXTENSION_EMF, False, False
        
'      ' Highlight each door in the sheet
'      Set LayerHighlight = ActiveDrawing.CreateLayer("HIGHLIGHT")
'      With LayerHighlight
'        .LineWidth = 9
'      End With
'      Set LayerAPS = ActiveDrawing.Layers(1)
'
'      lngSheetPartCount = ActiveDrawing.Attribute(DEF_ATT_SHEET_DOOR_COUNT)
'
'      For lngPartCount = 1 To lngSheetPartCount
'        For Each SheetPath In ActiveDrawing.Geometries
'          strPartCount = SheetPath.Attribute(DEF_ATT_NEST_DOOR_COUNT)
'          If strPartCount <> "" Then
'            If strPartCount = CStr(lngPartCount) Then
'              SheetPath.SetLayer LayerHighlight
'              SheetPath.Color = acamRED
'            Else
'              SheetPath.SetLayer LayerAPS
'              SheetPath.Color = acamLIGHT_GREY
'            End If
'            SheetPath.Redraw
'          End If
'        Next
'
'        sEMFPath = gstr_CheckDir(Frame.PathOfThisAddin & DEF_BACKSLASH & DEF_PATH_IMAGE) & _
'          gstr_JobName & DEF_UNDERSCORE & PressDetails & DEF_UNDERSCORE & colSheetNames(iCount) & DEF_UNDERSCORE & lngPartCount
'
'        ActiveDrawing.SaveEmfFile sEMFPath & DEF_EXTENSION_EMF, False, False
'
'      Next
    Next
    
    ' Restore the background to its original colour
    ActiveDrawing.BackgroundColorWireGradient = blnGradientFillWire
    ActiveDrawing.BackgroundColorShadingGradient = blnGradientFillShaded
    ActiveDrawing.BackgroundColor = lBackgroundColour
'    ActiveDrawing.RedrawShadedViews

    Frame.CloseProgressBox

    ' Restore original nested drawing
    With clsOptions
        strNestARD = mstr_CompileNestOrNCFilenamePress(PressDetails) & DEF_EXTENSION_ARD
        If .OutputResultsSubFolder Then
          OpenDrawing gstr_EnsureBackslash(.PathToRoot) & gstr_JobName & "\" & strNestARD
        Else
          OpenDrawing gstr_EnsureBackslash(.PathToRoot) & strNestARD
        End If
    End With
    
End Sub

Public Sub m_CreateAlphaCAMDrawingsOfSheets(Material As CMaterial)

    Dim nInfo                   As NestInformation
    Dim SH                      As NestSheet
    Dim SheetPath               As Path
    Dim DrawingText             As Text
    Dim SheetDrawing            As Drawing
    Dim strNestImage            As String
    Dim strNestARD              As String
    Dim TextPaths               As Paths
    Dim TextPath                As Path
    Dim lBackgroundColour       As Long
    Dim iNumSheets              As Integer
    Dim iCount                  As Integer
    Dim sEMFPath                As String
    Dim colSheetNames           As New Collection
    Dim SheetIdent              As String
    Dim Md                      As MillData
    Dim Npi                     As NestPartInstance
    Dim NestPath                As Path
    Dim lngPartCount            As Long
    Dim LayerHighlight          As Layer
    Dim LayerAPS                As Layer
    Dim strPartCount            As String
    Dim lngSheetPartCount       As Long
    Dim SheetMinX               As Double
    Dim SheetMinY               As Double
    Dim SheetMaxX               As Double
    Dim SheetMaxY               As Double
    Dim strSave                 As String
    Dim blnQuickShade           As Boolean
    Dim pthNestZone             As Path
    Dim SheetPath2              As Path
    Dim ps                      As Paths
'
    Set nInfo = ActiveDrawing.GetNestInformation
    ActiveDrawing.ThreeDViews = False
    ActiveDrawing.QuickShading = False

    ActiveDrawing.ScreenUpdating = False
    Frame.ProjectBarUpdating = False
    App.DisableUndo = True

    iNumSheets = nInfo.Sheets.Count
    
    ' Hide nest zones in reports
    For Each pthNestZone In pthsNestZones
      pthNestZone.Visible = False
    Next
    
    With nInfo
        For Each SH In nInfo.Sheets
            
            lngPartCount = 1
            
            For Each Npi In SH.Parts
              
              strNestImage = gstr_CheckDir(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) & _
                                           gstr_JobName & DEF_UNDERSCORE & Material.MaterialName & DEF_UNDERSCORE & SH.Name & DEF_UNDERSCORE & lngPartCount
              
              For Each NestPath In Npi.Paths
                NestPath.Attribute(DEF_ATT_NEST_DOOR_IMAGE) = strNestImage
                NestPath.Attribute(DEF_ATT_NEST_DOOR_COUNT) = lngPartCount
              Next
              
              lngPartCount = lngPartCount + 1
              
            Next
                    
        Next
        
    End With

    With clsOptions
        strNestARD = mstr_CompileNestOrNCFilename(Material) & DEF_EXTENSION_ARD
        'strNestARD = DEF_NEST_PREFIX & udtCCI.JobID & DEF_UNDERSCORE & clsNest.SheetName & DEF_EXTENSION_ARD
        Frame.ShowProgressBox Frame.ReadTextFile(strCTX, 300, 22), Frame.ReadTextFile(strCTX, 300, 23)

        If clsOptions.OutputResultsSubFolder Then
          strSave = gstr_EnsureBackslash(.PathToRoot) & gstr_JobName & "\" & strNestARD
        Else
          strSave = gstr_EnsureBackslash(.PathToRoot) & strNestARD
        End If

        ActiveDrawing.SaveAs strSave
        
        ' TFS #56774
        If Not clsOptions.SaveAllNestARD Then
          colDeleteFiles.add strSave
        End If
        
    End With
    

    ' Convert all text on the active drawing to geometry
    For Each DrawingText In ActiveDrawing.Text
        SheetIdent = DrawingText.Attribute(attSheetIdent)
        Set TextPaths = DrawingText.ConvertToGeometry
        If SheetIdent <> "" Then
            For Each TextPath In TextPaths
                TextPath.Attribute(attSheetIdent) = SheetIdent
            Next
        End If
    Next
    
    Dim t1 As Single
    Dim T2 As Single

    t1 = Timer
        
    SuppressUpdateRapids True
    
    With nInfo
        For Each SH In nInfo.Sheets

            Set SheetDrawing = App.CreateTempDrawing
            SheetIdent = gstr_GetNestedSheetIdent(SH.Name)

            Frame.SetProgressText Frame.ReadTextFile(DEF_TEXT, 300, 13) & Space(1) & SH.Name

            For Each SheetPath In SH.Paths
                If SheetPath.IsToolPath Then
                    Set Md = SheetPath.GetMillData
                    If Md.ProcessType = 2 Or Md.ProcessType = 4 Or Md.ProcessType = 5 Then
                      SheetPath.Attribute(DEF_ATT_POCKET_PATH) = "1"
                    End If
                End If
                
                If PathHasLeadInOut(SheetPath) Then
                  If SheetPath.ToolInOut = acamON_CENTER Then
                    SheetPath.SetLeadInOutManual acamLeadNONE, acamLeadNONE, False, False, 0, 0, 0, 0
                  Else
                    SheetPath.SetLeadInOutAuto acamLeadNONE, acamLeadNONE, 1, 1, 0, False, False, 0
                  End If
                End If
                
                SheetPath.SetLayer ActiveDrawing.Layers(1)
                SheetPath.MoveToDrawing SheetDrawing
            Next

            For Each TextPath In ActiveDrawing.Geometries

                If TextPath.TestInsidePath(SH.Path) = acamResultTRUE _
                  Or TextPath.Attribute(attSheetIdent) = SheetIdent Then

                    TextPath.SetLayer ActiveDrawing.Layers(1)
                    TextPath.MoveToDrawing SheetDrawing

                End If

            Next


            SH.Path.Sheet = False
            SH.Path.MoveToDrawing SheetDrawing

            ' Add the sheet name to a local collection (for language inconsistencies)
            colSheetNames.add SH.Name

            SheetDrawing.Attribute(DEF_ATT_SHEET_DOOR_COUNT) = SH.Parts.Count

            strNestImage = gstr_CheckDir(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) & _
                                         gstr_JobName & DEF_UNDERSCORE & Material.MaterialName & DEF_UNDERSCORE & SH.Name


            SheetDrawing.SaveAs strNestImage & DEF_EXTENSION_ARD

        Next
    End With

    T2 = Timer

''    MsgBox Format(T2 - t1, "0.00") & " secs"

    Frame.SetProgressText Frame.ReadTextFile(strCTX, 300, 24)

    ' ljo - 15/7/11
    ' Prevents invisible report images with new Alphacam shading modes
    ' TFS#45271

    Dim blnGradientFillWire     As Boolean
    Dim blnGradientFillShaded   As Boolean

    ' Save the current background colour
    blnGradientFillWire = ActiveDrawing.BackgroundColorWireGradient
    blnGradientFillShaded = ActiveDrawing.BackgroundColorShadingGradient
    
    ActiveDrawing.BackgroundColorShadingGradient = False
    ActiveDrawing.BackgroundColorWireGradient = False
    lBackgroundColour = ActiveDrawing.BackgroundColor
    ActiveDrawing.BackgroundColor = acamWHITE
'    ActiveDrawing.RedrawShadedViews

    ' Save the current background colour
    'lBackgroundColour = ActiveDrawing.BackgroundColor
    'ActiveDrawing.BackgroundColor = acamWHITE

    For iCount = 1 To iNumSheets

      ' Retrieve the AlphaCAM drawing using the name stored in the local collection
      strNestARD = gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE & DEF_BACKSLASH & _
        gstr_JobName & DEF_UNDERSCORE & Material.MaterialName & DEF_UNDERSCORE & colSheetNames(iCount)

      App.OpenDrawing strNestARD & DEF_EXTENSION_ARD

      For Each SheetPath In ActiveDrawing.Geometries
        If SheetPath.Attribute(DEF_ATT_POCKET_PATH) = "1" Then
          SheetPath.Color = acamYELLOW
        Else
          SheetPath.Color = acamBLACK
        End If
'        SheetPath.Redraw
      Next

      sEMFPath = gstr_CheckDir(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) & _
        gstr_JobName & DEF_UNDERSCORE & Material.MaterialName & DEF_UNDERSCORE & colSheetNames(iCount)

      ' 04/20/06 - rg
      '     + ADDED so that entire drawing is present in report, this
      '             was not an issue until I added screen locking
      '
      'Call ActiveDrawing.AutoScale  ' .ZoomAll

      ' Zoom in based on the sheet extents
      ActiveDrawing.GetExtent SheetMinX, SheetMinY, 0, SheetMaxX, SheetMaxY, 0

      'ActiveDrawing.ScreenUpdating = True

      ActiveDrawing.ZoomToBox SheetMinX, SheetMinY, SheetMaxX, SheetMaxY, 1

      'ActiveDrawing.ScreenUpdating = False

      g_ShowHideScrapCuts False
      
      ActiveDrawing.SaveEmfFile sEMFPath & DEF_EXTENSION_EMF, False, False
'EditMark
     ' Highlight each door in the sheet
    Set LayerHighlight = ActiveDrawing.CreateLayer("HIGHLIGHT")
    With LayerHighlight
    .LineWidth = 9
    End With
    Set LayerAPS = ActiveDrawing.Layers(1)
    lngSheetPartCount = ActiveDrawing.Attribute(DEF_ATT_SHEET_DOOR_COUNT)
      For lngPartCount = 1 To lngSheetPartCount
        Set SheetPath2 = Nothing
        For Each SheetPath In ActiveDrawing.Geometries
          strPartCount = SheetPath.Attribute(DEF_ATT_NEST_DOOR_COUNT)
          If strPartCount <> "" Then
            If strPartCount = CStr(lngPartCount) Then

                If Not SheetPath2 Is Nothing Then
                    If SheetPath.GetArea(-1) >= SheetPath2.GetArea(-1) Then
                        SheetPath2.Delete
                        Set SheetPath2 = SheetPath
                    Else
                        SheetPath.Delete
                    End If
                Else
                    Set SheetPath2 = SheetPath
                End If
            End If
            SheetPath.Redraw
          End If
        Next
      Next

For lngPartCount = 1 To lngSheetPartCount

        For Each SheetPath In ActiveDrawing.Geometries

          strPartCount = SheetPath.Attribute(DEF_ATT_NEST_DOOR_COUNT)

          If strPartCount <> "" Then
            If strPartCount = CStr(lngPartCount) Then
              SheetPath.SetLayer LayerHighlight
              SheetPath.Color = acamRED

             Set ps = App.ActiveDrawing.HatchPath(SheetPath, acamHatchSingle, 45, 5, 10)


            Else
              SheetPath.SetLayer LayerAPS
              SheetPath.Color = acamLIGHT_GREY
            End If
            SheetPath.Redraw
          Else
            If SheetPath.Attribute(attScrapCut) = "1" Then
              SheetPath.Visible = False
              SheetPath.Redraw
            End If
          End If

        Next

        sEMFPath = gstr_CheckDir(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) & _
          gstr_JobName & DEF_UNDERSCORE & Material.MaterialName & DEF_UNDERSCORE & colSheetNames(iCount) & DEF_UNDERSCORE & lngPartCount
        ActiveDrawing.SaveEmfFile sEMFPath & DEF_EXTENSION_EMF, False, False

        ps.Delete

      Next
      Next
      
'       Set LayerHighlight = ActiveDrawing.CreateLayer("HIGHLIGHT")
'      With LayerHighlight
'        .LineWidth = 9
'      End With
'      Set LayerAPS = ActiveDrawing.Layers(1)
'
'      lngSheetPartCount = ActiveDrawing.Attribute(DEF_ATT_SHEET_DOOR_COUNT)
'
'      For lngPartCount = 1 To lngSheetPartCount
'        For Each SheetPath In ActiveDrawing.Geometries
'          strPartCount = SheetPath.Attribute(DEF_ATT_NEST_DOOR_COUNT)
'          If strPartCount <> "" Then
'            If strPartCount = CStr(lngPartCount) Then
'              SheetPath.SetLayer LayerHighlight
'              SheetPath.Color = acamRED
'            Else
'              SheetPath.SetLayer LayerAPS
'              SheetPath.Color = acamLIGHT_GREY
'            End If
'            SheetPath.Redraw
'          Else
'            If SheetPath.Attribute(attScrapCut) = "1" Then
'              SheetPath.Visible = False
'              SheetPath.Redraw
'            End If
'          End If
'        Next
'
'        sEMFPath = gstr_CheckDir(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) & _
'          gstr_JobName & DEF_UNDERSCORE & Material.MaterialName & DEF_UNDERSCORE & colSheetNames(iCount) & DEF_UNDERSCORE & lngPartCount
'
'        ActiveDrawing.SaveEmfFile sEMFPath & DEF_EXTENSION_EMF, False, False
'
'      Next
'


    ' Restore the background to its original colour
    ActiveDrawing.BackgroundColorWireGradient = blnGradientFillWire
    ActiveDrawing.BackgroundColorShadingGradient = blnGradientFillShaded
    ActiveDrawing.BackgroundColor = lBackgroundColour
'    ActiveDrawing.RedrawShadedViews

    Frame.CloseProgressBox

    'ActiveDrawing.ScreenUpdating = True
    'Frame.ProjectBarUpdating = True
    App.DisableUndo = False

    ActiveDrawing.QuickShading = blnQuickShade

    ' Restore original nested drawing
    With clsOptions
        strNestARD = mstr_CompileNestOrNCFilename(Material) & DEF_EXTENSION_ARD
        'strNestARD = DEF_NEST_PREFIX & udtCCI.JobID & DEF_UNDERSCORE & clsNest.SheetName & DEF_EXTENSION_ARD
        OpenDrawing strSave
    End With

    SuppressUpdateRapids False

End Sub

Private Function mbln_ComparePressNestToOrder(Press As CPress, PressColour As CPressColour, Optional CurrentPressThickness As Double) As Boolean
  Dim Ni                As NestInformation
  Dim Sheet             As NestSheet
  Dim lQuantity         As Long
  Dim NestPartCount     As Long
  Dim PressThickness    As CPressThickness
  Dim Door              As CDoor
'
  lQuantity = 0
  
  If clsOptions.GroupByMaterialThickness Then
    For Each Door In Press.colPressColours(PressColour.ColourName).colPressThicknesses("k" & CurrentPressThickness).colPressComponents
      lQuantity = lQuantity + Door.Quantity
    Next
  Else
    For Each PressThickness In PressColour.colPressThicknesses
      For Each Door In Press.colPressColours(PressColour.ColourName).colPressThicknesses("k" & PressThickness.PressThickness).colPressComponents
        lQuantity = lQuantity + Door.Quantity
      Next
    Next
  End If
    
  Set Ni = ActiveDrawing.GetNestInformation
        
  For Each Sheet In Ni.Sheets
    NestPartCount = NestPartCount + Sheet.Parts.Count
  Next Sheet
    
  If NestPartCount = lQuantity Then
    mbln_ComparePressNestToOrder = True
  End If
End Function

Private Function mbln_ProcessPart(Door As CDoor, lngPartNumber As Long, _
                                  colANC As Collection, bPress As Boolean, bPreview As Boolean) As Boolean
    
    Dim bSuccess                As Boolean
    Dim bStopProcessing         As Boolean
    'Dim sCTX                    As String
    
On Error GoTo mbln_ProcessPart_Error
    
    '..get the style number - must add case for every door style
    Select Case Door.StyleNumber

        Case 900    '..inserted ARD

            If Not mbln_Style_900(Door, lngPartNumber, colANC, bPreview, bPress, bStopProcessing) Then                   '..inserted ARD
                 If bStopProcessing Then
                    
                    bSuccess = False
                    GoTo Controlled_Exit
                 
                 Else
                    MsgBox Frame.ReadTextFile(strCTX, 500, 97) & Space(1) & UCase$(Door.TypeName), vbExclamation, DEF_PROJECT_NAME
                    
                    '..set successful so that we don't get another message later
                    bSuccess = True
                 
                 End If
            
            Else
                
                bSuccess = True
                
            End If

        Case 910    '..inserted ARM

            If Not mbln_Style_910(Door, lngPartNumber, colANC, bPreview) Then                     '..inserted ARD
                 MsgBox Frame.ReadTextFile(strCTX, 500, 97) & Space(1) & UCase$(Door.TypeName), vbExclamation, DEF_PROJECT_NAME
            End If

            '..set successful so that we don't get another message later
            bSuccess = True

        Case 920    '..inserted ARB

            If Not mbln_Style_920(Door, lngPartNumber, colANC, bPreview) Then                   '..inserted ARD
                 MsgBox Frame.ReadTextFile(strCTX, 500, 97) & Space(1) & UCase$(Door.TypeName), vbExclamation, DEF_PROJECT_NAME
            End If

            '..set successful so that we don't get another message later
            bSuccess = True

        Case 930: bSuccess = mbln_Style_Make_930(Door, bPress, lngPartNumber, colANC, bPreview)         '..user defined geometry
                                          
        Case Else                                                                                                                   '..unknown
            
            With Frame
            
                '..look for preview and if not then give option to continue
                If Not bPreview Then
            
                    '..tell the user and ask to continue
                    If MsgBox(.ReadTextFile(strCTX, 500, 21) & Space(1) & Door.StyleNumber & Space(3) & vbCrLf & vbCrLf & _
                              .ReadTextFile(strCTX, 500, 20), vbExclamation + vbYesNo, DEF_PROJECT_NAME) = vbNo Then
                                  
                        bSuccess = False
                        GoTo Controlled_Exit
                        
                    End If
                    
                Else
                    
                    bSuccess = False
                
                End If

            End With
                                                                                           
    End Select
    
    '..did it work?
    If Not bSuccess Then
        
        App.New
        
        '..look for preview and if not then give option to continue
        If Not bPreview Then
                    
            With Door
            
                '..tell the user and ask to continue
                If MsgBox(Frame.ReadTextFile(strCTX, 500, 18) & Space(1) & .StyleNumber & _
                          DEF_COMMA & Frame.ReadTextFile(strCTX, 500, 19) & Space(1) & .TypeName & vbCrLf & Frame.ReadTextFile(strCTX, 500, 20), _
                          vbExclamation + vbYesNo, Frame.ReadTextFile(strCTX, 120, 1)) = vbYes Then
                          
                          bSuccess = True
                          
                End If
                
            End With
            
        End If
        
    End If
    
Controlled_Exit:
    
    mbln_ProcessPart = bSuccess

Exit Function

mbln_ProcessPart_Error:
    
    bSuccess = False
    Resume Controlled_Exit

End Function

Private Sub m_FillDoorData(Door As CDoor, rDetails As ADODB.Recordset, bPreview As Boolean)

    With Door
        
        .DetailID = gvar_CheckNull(rDetails.Fields!PK)
        
        If bPreview Then
            
            .CornerRadius = gvar_CheckNull(rDetails.Fields!CornerRadius)
            .Quantity = 1
            .StyleNumber = 930
            .TypeName = gvar_CheckNull(rDetails.Fields!TypeID)
            .Width = gvar_CheckNull(rDetails.Fields!Width)
            .Length = gvar_CheckNull(rDetails.Fields!Length)
            .IgnoreOuterGeometry = gvar_CheckNull(rDetails.Fields!IgnoreOuterGeometry)
            .UserStyleName = gvar_CheckNull(rDetails.Fields!UserStyleName)
                
        Else
        
            .OrderID = gvar_CheckNull(rDetails.Fields!OrderID)
            .CornerRadius = gvar_CheckNull(rDetails.Fields!CornerRadius)
            .Length = gvar_CheckNull(rDetails.Fields!Length)
            .Quantity = gvar_CheckNull(rDetails.Fields!Quantity)
            .StyleNumber = gvar_CheckNull(rDetails.Fields!StyleNumber)
            .TypeName = gvar_CheckNull(rDetails.Fields!TypeName)
            .Width = gvar_CheckNull(rDetails.Fields!Width)
            .UserStyleName = gvar_CheckNull(rDetails.Fields!StyleName)
            
            .IgnoreOuterGeometry = gvar_CheckNull(rDetails.Fields!IgnoreOuterGeometry)
            .RotationMethod = gvar_CheckNull(rDetails.Fields!RotationMethod)
            .RotationAngle = gvar_CheckNull(rDetails.Fields!RotationAngle)
            .ColourRotationMethod = gvar_CheckNull(rDetails.Fields!ColourRotationMethod)
            .ColourDefinedRotationMethod = gvar_CheckNull(rDetails.Fields!ColourRot)
            
            .ByPassNest = gvar_CheckNull(rDetails.Fields!ByPassNest)                                            ' reinstated 07/28/03 - rg
            .OversizeX = gvar_CheckNull(rDetails.Fields!OversizeX)                                              ' 07/28/03 - rg
            .OversizeY = gvar_CheckNull(rDetails.Fields!OversizeY)                                              ' 07/28/03 - rg
            
            .NestingPriority = gvar_CheckNull(rDetails.Fields!NestingPriority)
            .NestingZone = gvar_CheckNull(rDetails.Fields!ZoneNumber)
            .NestingZoneID = gvar_CheckNull(rDetails.Fields!NestZoneID)
                                    
            .CSV_CustomerName = gvar_CheckNull(rDetails.Fields!CSV_CustomerName)
            .CSV_OrderNumber = gvar_CheckNull(rDetails.Fields!CSV_OrderNumber)
            .CSV_ItemNumber = gvar_CheckNull(rDetails.Fields!CSV_ItemNumber)
            .ProductionComment = gvar_CheckNull(rDetails.Fields!ProductionComment)
            .CustomField1 = gvar_CheckNull(rDetails.Fields!CustomField1)
            .CustomField2 = gvar_CheckNull(rDetails.Fields!CustomField2)
            .ComponentGrouping = gvar_CheckNull(rDetails.Fields!ComponentGrouping)
            
            .FoilColour = gvar_CheckNull(rDetails.Fields!ColourName)
            
        End If

        '..these are the same in both tables
        .UserVariableString = gvar_CheckNull(rDetails.Fields!UserVariableString)
        .UserVariableDescriptionString = gvar_CheckNull(rDetails.Fields!UserDescriptionString)
        .UserArg_0 = gvar_CheckNull(rDetails.Fields!UserValue_0)
        .UserArg_1 = gvar_CheckNull(rDetails.Fields!UserValue_1)
        .UserArg_2 = gvar_CheckNull(rDetails.Fields!UserValue_2)
        .UserArg_3 = gvar_CheckNull(rDetails.Fields!UserValue_3)
        .UserArg_4 = gvar_CheckNull(rDetails.Fields!UserValue_4)
        .UserArg_5 = gvar_CheckNull(rDetails.Fields!UserValue_5)
        .UserArg_6 = gvar_CheckNull(rDetails.Fields!UserValue_6)
        
        .HandleID = gvar_CheckNull(rDetails.Fields!HandleID)

    End With

End Sub

Private Sub m_FillMaterialData(Material As CMaterial, rstMaterial As ADODB.Recordset)
  
  With Material
    .MaterialName = gvar_CheckNull(rstMaterial.Fields!Name)
    
    ' TFS#80163
    .AbsolutePosition = glng_MaterialIndex
    glng_MaterialIndex = glng_MaterialIndex + 1
    
    .Width = gvar_CheckNull(rstMaterial.Fields!Width)
    .Length = gvar_CheckNull(rstMaterial.Fields!Length)
    .Thickness = gvar_CheckNull(rstMaterial.Fields!Thickness)
    .NoRotation = gvar_CheckNull(rstMaterial.Fields!NoRotation)
    .PackTo = gvar_CheckNull(rstMaterial.Fields!PackTo)
    .SearchRes = gvar_CheckNull(rstMaterial.Fields!SearchRes)
    .MinGapBetweenPaths = gvar_CheckNull(rstMaterial.Fields!MinGapBetweenPaths)
    .MinGapAtSheetEdge = gvar_CheckNull(rstMaterial.Fields!MinGapAtSheepEdge)
    .LeadGap = gvar_CheckNull(rstMaterial.Fields!LeadGap)
    .FinalZTolerance = gvar_CheckNull(rstMaterial.Fields!FinalZTolerance)
    .InsertFiller = gvar_CheckNull(rstMaterial.Fields!InsertFiller)
    .InsertFillerFile = gvar_CheckNull(rstMaterial.Fields!InsertFillerFile)
    .InsertFillerFile2 = gvar_CheckNull(rstMaterial.Fields!InsertFillerFile2)
    .InsertFillerFile3 = gvar_CheckNull(rstMaterial.Fields!InsertFillerFile3)
    .MinimizeToolChanges = gvar_CheckNull(rstMaterial.Fields!MinimizeToolChanges)
    .ToolPathsOnly = gvar_CheckNull(rstMaterial.Fields!ToolPathsOnly)
    .CutSmallPartsFirst = gvar_CheckNull(rstMaterial.Fields!CutSmallPartsFirst)
    .CutInnerPathsFirst = gvar_CheckNull(rstMaterial.Fields!CutInnerPathsFirst)
    .NCSubroutines = gvar_CheckNull(rstMaterial.Fields!NCSubroutines)
    .RotationFlip = gvar_CheckNull(rstMaterial.Fields!RotationFlip)
    .LeaveEdgeGapUncut = gvar_CheckNull(rstMaterial.Fields!LeaveEdgeGapUncut)
    .NumberComponentsBySize = gvar_CheckNull(rstMaterial.Fields!NumberComponentsBySize)
    .SuppressFinalSort = gvar_CheckNull(rstMaterial.Fields!SuppressFinalSort)
    .NestingScreenUpdate = gvar_CheckNull(rstMaterial.Fields!NestingScreenUpdate)
    .OnionSkin = gvar_CheckNull(rstMaterial.Fields!OnionSkin)
    .OnionSkinMinXY = gvar_CheckNull(rstMaterial.Fields!OnionSkinMinXY)
    .OnionSkinMinArea = gvar_CheckNull(rstMaterial.Fields!OnionSkinMinArea)
    .OnionSkinThickness = gvar_CheckNull(rstMaterial.Fields!OnionSkinThickness)
    .OnionSkinCutOrder = gvar_CheckNull(rstMaterial.Fields!OnionSkinCutOrder)
    .OnionSkinApplyToInside = gvar_CheckNull(rstMaterial.Fields!OnionSkinApplyToInside)
    .ProcessWaste = gvar_CheckNull(rstMaterial.Fields!ProcessWaste)
    .ProcessWasteMCStyle = gvar_CheckNull(rstMaterial.Fields!ProcessWasteMCStyle)
    .ProcessWasteDepthOfCut = gvar_CheckNull(rstMaterial.Fields!ProcessWasteDepthOfCut)
    .ProcessWasteFinalSheetScrap = gvar_CheckNull(rstMaterial.Fields!ProcessWasteFinalSheetScrap)
    .ProcessWasteCutTowardsComponents = gvar_CheckNull(rstMaterial.Fields!ProcessWasteCutTowardsComponents)
    .ProcessWasteStrategy = gvar_CheckNull(rstMaterial.Fields!ProcessWasteStrategy)
    .HorizontalSpacing = gvar_CheckNull(rstMaterial.Fields!HorizontalCutSpacing)
    .VerticalSpacing = gvar_CheckNull(rstMaterial.Fields!VerticalCutSpacing)
    .PackFinalSheetComponentsToLHS = gvar_CheckNull(rstMaterial.Fields!PackFinalSheetComponentsToLHS)
    .TimePerSheet = gvar_CheckNull(rstMaterial.Fields!TimePerSheet)
  End With

End Sub



Private Sub m_DeleteReportData(ByVal lOrder As Long, sJob As String)
        
    Dim rstJobName              As ADODB.Recordset
    Dim FSO                     As New Scripting.FileSystemObject
        
On Error Resume Next
    
    Set rstJobName = grst_GetJobName(lOrder)
    
    If Not (rstJobName Is Nothing) Then
        
        '..get job name - also returned to g_MakeMaster
        sJob = gvar_CheckNull(rstJobName.Fields!JobName)
        
        '..wipe out the images
        With FSO
            
            .DeleteFile gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE & DEF_BACKSLASH & sJob & DEF_UNDERSCORE & _
                           DEF_WILDCARD & DEF_EXTENSION_EMF, True
                               
        End With
    
        If (rstJobName.State = adStateOpen) Then rstJobName.Close
        Set rstJobName = Nothing
    
    End If
    
    If (Err.Number <> 0) Then Err.Clear
    
    DoEvents

    '..ensure connection
    If Not gbln_ConnectToDB Then Exit Sub

    gdb_CDM.Execute "DELETE FROM AD_REPORT_DATA WHERE OrderID=" & lOrder & ";"
    
    If (Err.Number <> 0) Then WriteError Err, False, "m_DeleteReportData": Err.Clear
    
End Sub

Private Function mstr_CurrentPost() As String

On Error Resume Next

    '..get the current post so we can change back later
    mstr_CurrentPost = App.PostFileName

End Function

Private Sub m_OutputToAlphaEDIT(colNestANC As Collection, colDOORANC As Collection)

    Dim aEDIT                   As AlphaEdit.App
    Dim iFileCount              As Integer
    Dim iU                      As Integer

On Error Resume Next
        
'    '..are we nesting?
'    If blnNesting Then
    
        '..check for any nested anc files
        If Not colNestANC.Count = 0 Then
            
            '..start up alphaedit
            Set aEDIT = New AlphaEdit.App
            
            iU = colNestANC.Count
            
            '..loop thru and open all nested anc files in alphaedit
            For iFileCount = 1 To iU
                aEDIT.OpenDoc colNestANC(iFileCount)
            Next iFileCount
            
        End If
    
'    Else    '..not nesting
    
        '..check for any single door anc files
        If Not colDOORANC.Count = 0 Then
        
            '..start alphaedit
            If (aEDIT Is Nothing) Then Set aEDIT = New AlphaEdit.App
        
            iU = colDOORANC.Count
            
            '..loop thru and open all single door anc files in alphaedit
            For iFileCount = 1 To iU
                aEDIT.OpenDoc colDOORANC(iFileCount)
            Next iFileCount
            
        End If
            
'    End If
    
    Set aEDIT = Nothing
    
End Sub

Private Sub m_UpdateDBNestPaths(sFieldName As String, sFieldPath As String, _
                                sPath As String, sName As String, bSplit As Boolean, Optional iSheet As Integer)
    
    Dim rst                     As ADODB.Recordset
    Dim strJob()                As String
    Dim intCount                As Integer
    Dim lJob                    As Long
    
On Error Resume Next
    
    '..make sure we're connected
    If Not gbln_ConnectToDB Then Exit Sub
                  
    strJob = Split(gstr_JobIDs, ",")
    
    For intCount = LBound(strJob) To UBound(strJob)
    
        lJob = strJob(intCount)
        
        '..get the records
        If bSplit Then
            Set rst = grst_GetReportDataNestSheet(lJob, clsNest.SheetName, iSheet)
        Else
            Set rst = grst_GetReportDataNest(lJob, clsNest.SheetName)
        End If
        
        If (rst Is Nothing) Then GoTo Controlled_Exit
        
        With rst
        
            '..get the first record
            .MoveFirst
    
            '..loop thru them all
            Do Until .EOF
                
                '..edit the anc file and path
                .Fields(sFieldPath) = sPath
                .Fields(sFieldName) = sName
                .Update
                
                '..get the next record
                .MoveNext
                
            Loop
        
            '..close it
            .Close
            
        End With
    
    Next
    
Controlled_Exit:
    
    Set rst = Nothing
    
    If (Err.Number <> 0) Then WriteError Err, False, "m_UpdateDBNestPaths"
    
Exit Sub
        
End Sub

Private Function mbln_InsertARD(Door As CDoor, bPreview As Boolean) As Boolean
    
    Dim FSO         As New Scripting.FileSystemObject
    Dim drwTemp     As Drawing
    Dim MinX        As Double
    Dim MinY        As Double
    Dim minZ        As Double
    Dim MaxX        As Double
    Dim MaxY        As Double
    Dim maxZ        As Double
    Dim sTmp        As String
    Dim bOK         As Boolean
    Dim pVol        As Path
    Dim lngTpPre    As Long
    Dim lngTpPost   As Long
    Dim lngCount    As Long
    Dim dLength     As Double
    Dim dWidth      As Double
    
On Error GoTo mbln_InsertARD_Error
        
    '..start out fine
    mbln_InsertARD = True
    bOK = True
    
    dLength = Door.Length
    dWidth = Door.Width
    
    
    'sTmp = App.Frame.PathOfThisAddin & DEF_BACKSLASH & "tmp.ard"
    sTmp = gs_GetCommonAppDataDir & "tmp.ard"
    
    '..insert it
    With clsPathData
    
        '..make sure the insert file exists
        If Not FSO.FileExists(.InsertFilePath) Then mbln_InsertARD = False: GoTo Controlled_Exit
        
        '..store grouping info in attributes for parametric inserted drawings
        ' (saving destroys custom group numbers)
        g_ConvertGroupsToAttributes
        
        '..save the current drawing to a temp file
        ActiveDrawing.SaveAs sTmp
                
        '..open the file to insert and validate it
        App.OpenDrawing .InsertFilePath
        
        '..validate the file
        Set drwTemp = App.ActiveDrawing '  .OpenTempDrawing(.InsertFilePath)
        
        '..look for workplanes
        If (drwTemp.WorkPlanes.Count <> 0) Then
            MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 600, 146) & Space(3), vbInformation, DEF_PROJECT_NAME
            bOK = False
        End If
        
        ' Only check depths if component is being manufactured
        ' Material information will not be available in preview mode
        ' so cannot check depths of cut etc.
        If Not bPreview Then
            '..loop thru all toolpaths looking for invalid depth
            For Each pVol In drwTemp.ToolPaths
                
                With pVol
                    bOK = gbln_CheckDepthTolerance(.GetMillData.MaterialTop, .GetMillData.FinalDepth)
                End With
            
            Next
        End If
        
        '..get the extents of the file to be inserted
        
        ' TFS#81512 - Changes made by TFS#59874 below will break existing customers setups
        ' This is fixed by re-instating the existing extents method (drawing extents rather
        ' than toolpath extents) - there is now an option to select to use the existing method
        If clsOptions.UseDrawingExtentsForInsertedDrawings Then
          drwTemp.GetExtent MinX, MinY, minZ, MaxX, MaxY, maxZ
        Else
          ' TFS #59874 - Drawing extents includes tool diameter, causing insert failure
          ' when using large diameter tools
          drwTemp.ToolPaths.GetExtentL MinX, MinY, MaxX, MaxY
        End If
        
        '..check insert dimensions against current door dimensions to make sure it will fit
        If ((MaxX - MinX) > dWidth) Then bOK = False
        If ((MaxY - MinY) > dLength) Then bOK = False
        
        '..reopen the original drawing
        App.OpenDrawing sTmp
        
        '..re-apply custom group numbers (from attributes)
        g_ConvertAttributesToGroups
        
        '..everything ok so far?
        If Not bOK Then mbln_InsertARD = False: GoTo Controlled_Exit
        
        lngTpPre = ActiveDrawing.GetToolPathCount
        
        '..find the reference point
        Select Case .InsertFileReferencePoint
            
            Case adoorINSERT_AUTO_CENTER: ActiveDrawing.InsertDrawing .InsertFilePath, (dWidth / 2), (dLength / 2), 0
            Case adoorINSERT_CENTER_BOTTOM_X: ActiveDrawing.InsertDrawing .InsertFilePath, (dWidth / 2), (MinY + .InsertFilePointY), 0
            Case adoorINSERT_CENTER_TOP_X: ActiveDrawing.InsertDrawing .InsertFilePath, (dWidth / 2), (MaxY - .InsertFilePointY), 0
            Case adoorINSERT_CENTER_LEFT_Y: ActiveDrawing.InsertDrawing .InsertFilePath, (MinX + .InsertFilePointX), (dLength / 2), 0
            Case adoorINSERT_CENTER_RIGHT_Y: ActiveDrawing.InsertDrawing .InsertFilePath, (MaxX - .InsertFilePointX), (dLength / 2), 0
            Case adoorINSERT_TOP_LEFT: ActiveDrawing.InsertDrawing .InsertFilePath, .InsertFilePointX, (dLength - .InsertFilePointY), 0
            Case adoorINSERT_TOP_RIGHT: ActiveDrawing.InsertDrawing .InsertFilePath, (dWidth - .InsertFilePointX), (dLength - .InsertFilePointY), 0
            Case adoorINSERT_BOTTOM_LEFT: ActiveDrawing.InsertDrawing .InsertFilePath, .InsertFilePointX, .InsertFilePointY, 0
            Case adoorINSERT_BOTTOM_RIGHT: ActiveDrawing.InsertDrawing .InsertFilePath, (dWidth - .InsertFilePointX), .InsertFilePointY, 0
            Case adoorINSERT_PARAMETRIC
                If Not gbln_InsertDrawing(.InsertFilePath, .InsertFileGroupNumber) Then
                    mbln_InsertARD = False
                    GoTo Controlled_Exit
                End If
            Case Else: mbln_InsertARD = False: GoTo Controlled_Exit
            
        End Select
                
    End With
    
    lngTpPost = ActiveDrawing.GetToolPathCount
    
    For lngCount = (lngTpPre + 1) To lngTpPost
      m_SetAttributes ActiveDrawing.ToolPaths(lngCount), Door
    Next
    
    '..clean out any work volume if exists
    For Each pVol In ActiveDrawing.Geometries
         
        With pVol
            .Selected = .IsWorkVolume
        End With
        
    Next pVol
    
    '..delete the selected (should be the work volume)
    ActiveDrawing.DeleteSelected
        
Controlled_Exit:

    Set drwTemp = Nothing
    Set FSO = Nothing

Exit Function

mbln_InsertARD_Error:

    If (Err.Number <> 0) Then WriteError Err, True, "mbln_InsertARD"
    mbln_InsertARD = False
    Resume Controlled_Exit

End Function


Private Function mbln_Intialize(lngGeo As Long) As Boolean

On Error GoTo mbln_Intialize_Error

    '..start this function off on the right track
    mbln_Intialize = True
    
    '..start off the geo count at zero
    lngGeo = 0
        
    '..new drawing
    App.New
    
    '..clean up the screen
    'ActiveDrawing.Redraw
    
    DoEvents
    
Exit Function
    
mbln_Intialize_Error:

    If (Err.Number <> 0) Then WriteError Err, True, "mbln_Intialize"
    mbln_Intialize = False

End Function

Private Sub m_SetAttributes(Pth As Path, Door As CDoor)    ' strType As String, dWidth As Double, dLength As Double)   '..07.09.02 - rg
    
    Dim dblArea                 As Double
    
On Error Resume Next
    
    With Pth
        
        .Attribute(DEF_ATT_ALPHADOOR) = 1
        .Attribute(DEF_ATT_DETAIL_ID) = Door.DetailID
        .Attribute(DEF_ATT_STYLE_NUMBER) = Door.StyleNumber
        .Attribute(DEF_ATT_TYPE_NAME) = Door.TypeName
        .Attribute(DEF_ATT_TOOL_NAME) = gstr_StripLicomDatPath(clsPathData.ToolFullPath)
        .Attribute(DEF_ATT_PART_WIDTH) = IIf((Door.Width = 0), DEF_NOT_APPLICABLE, CStr(Door.Width))
        .Attribute(DEF_ATT_PART_LENGTH) = IIf((Door.Length = 0), DEF_NOT_APPLICABLE, CStr(Door.Length))
        dblArea = Door.Width * Door.Length
        .Attribute(DEF_ATT_PART_AREA) = IIf((Door.Width = 0), DEF_UNKNOWN, CStr(dblArea))
        .Attribute(DEF_ATT_GROUP_ID) = Pth.Group
        
        .Attribute(DEF_ATT_CUST_NAME) = Door.CSV_CustomerName
        .Attribute(DEF_ATT_ORDER_NUM) = Door.CSV_OrderNumber
        .Attribute(DEF_ATT_ITEM_NUM) = Door.CSV_ItemNumber
        .Attribute(DEF_ATT_ORDER_ID) = Door.OrderID
        .Attribute(DEF_ATT_DOOR_PRODUCTION_COMMENT) = Door.ProductionComment
        .Attribute(DEF_ATT_USER_STYLE_NAME) = Door.UserStyleName
        
        .Attribute(DEF_ATT_HANDLE_NAME) = gstr_GetHandleName(Door.HandleID)
        
        .Attribute(DEF_ATT_JOB_NAME) = Door.JobName
        .Attribute(DEF_ATT_COMPONENT_GROUPING) = Door.ComponentGrouping
                
        If Door.ByPassNest Then
            .Attribute(DEF_ATT_SHEET_THICKNESS) = clsNest.SheetThickness
            .Attribute(DEF_ATT_SHEET_LENGTH) = CStr(Door.Length + Door.OversizeY)
            .Attribute(DEF_ATT_SHEET_WIDTH) = CStr(Door.Width + Door.OversizeX)
        End If
                
    End With
    
    If (Err.Number <> 0) Then WriteError Err, False, "m_SetAttributes"

End Sub

Private Function mbln_FillPathData(rst As ADODB.Recordset) As Boolean

On Error GoTo mbln_FillPathData_Error
    
    '..start out fine
    mbln_FillPathData = True
    
    '..fill in the path data from the database
    With clsPathData
    
        .ApplyCompOnRapid = gvar_CheckNull(rst.Fields!CompOnRapid)
        .ChordError = gvar_CheckNull(rst.Fields!ChordError)
        .CutDirection = gvar_CheckNull(rst.Fields!CutDirection)
        .DepthOfCutsSpecified = gvar_CheckNull(rst.Fields!DepthsOfCutSpecified)
        
        ' Test to see if an engraving corner angle has been specified
        If IsNull(rst.Fields!EngraveCornerAngle) Then
          ' No Angle specified - so use the default value of 170?
          .EngraveCornerAngle = DEF_ENGRAVE_CORNER_ANGLE_LIMIT
        Else
          .EngraveCornerAngle = gvar_CheckNull(rst.Fields!EngraveCornerAngle)
        End If
        
        .FinalDepth = gvar_CheckNull(rst.Fields!FinalDepth)
        .FinalDepthPercentage = gvar_CheckNull(rst.Fields!FinalDepthPercentage)
        .IsFinalDepthPercent = gvar_CheckNull(rst.Fields!IsFinalDepthPercent)
        .FirstCutDepthPercentage = gvar_CheckNull(rst.Fields!ThicknessFirstCutPercent)
        .LastCutDepthPercentage = gvar_CheckNull(rst.Fields!ThicknessLastCutPercent)
        .FinalPassAroundIslands = gvar_CheckNull(rst.Fields!FinalPassIsIsland)
        .LastModified = gvar_CheckNull(rst.Fields!LastModified)
        .GroupID = gvar_CheckNull(rst.Fields!GroupID)                               '..04.11.02 - rg
        .LeadApproachAngle = gvar_CheckNull(rst.Fields!LeadApproachAngle)
        .LeadArcRadius = gvar_CheckNull(rst.Fields!LeadArcRadius)
        .LeadEntryPointIsCorner = gvar_CheckNull(rst.Fields!LeadEntryPointIsCorner)
        .LeadInSloping = gvar_CheckNull(rst.Fields!SlopeIn)
        .LeadInType = gvar_CheckNull(rst.Fields!LeadIn)
        .LeadLineLengthIn = gvar_CheckNull(rst.Fields!LeadLineLength)               '..07.18.02 - rg
        .LeadLineLengthOut = gvar_CheckNull(rst.Fields!LeadLineLengthOut)           '..07.18.02 - rg
        .LeadOutSloping = gvar_CheckNull(rst.Fields!SlopeOut)
        .LeadOutType = gvar_CheckNull(rst.Fields!LeadOut)
        .LeadOverlap = gvar_CheckNull(rst.Fields!LeadOverlap)
        .MachineCompensation = gvar_CheckNull(rst.Fields!McComp)
        .MachineMethod = gvar_CheckNull(rst.Fields!MachiningMethod)
        .MaterialTop = gvar_CheckNull(rst.Fields!MaterialTop)
        .MultiplePasses = gvar_CheckNull(rst.Fields!MultiplePasses)
        .NumberOfCuts = gvar_CheckNull(rst.Fields!NumberOfCuts)
        .OperationNumber = gvar_CheckNull(rst.Fields!PathNumber)
        .PathOffsetFrom = gvar_CheckNull(rst.Fields!PathOffsetFrom)
        .PathOffsetSide = gvar_CheckNull(rst.Fields!PathOffsetSide)
        .PathOffsetValue = gvar_CheckNull(rst.Fields!PathOffsetValue)
        .PocketBoundary = gvar_CheckNull(rst.Fields!PocketBoundary)
        .PocketType = gvar_CheckNull(rst.Fields!PocketType)
        .Pocket3DApproach = gvar_CheckNull(rst.Fields!Pocket3DApproach)
        .RapidDownTo = gvar_CheckNull(rst.Fields!RapidDownTo)
        .SafeRapidLevel = gvar_CheckNull(rst.Fields!SafeRapid)
        .StartCutting = gvar_CheckNull(rst.Fields!StartCutting)
        .StepLength = gvar_CheckNull(rst.Fields!StepLength)
        .Stock = gvar_CheckNull(rst.Fields!Stock)
        .ThicknessOfFirstCut = gvar_CheckNull(rst.Fields!ThicknessFirstCut)
        .ThicknessOfLastCut = gvar_CheckNull(rst.Fields!ThicknessLastCut)
        .ToolDiameter = gvar_CheckNull(rst.Fields!Diameter)
        .ToolDirectionCW = gvar_CheckNull(rst.Fields!ToolDirectionCW)
        .ToolDirectionIsReversed = gvar_CheckNull(rst.Fields!ToolDirectoinReversed)
        .ToolFullPath = LicomdatPath & (gvar_CheckNull(rst.Fields!ToolFullPath))
        .ToolInOut = gvar_CheckNull(rst.Fields!ToolInOut)
        .ToolName = gvar_CheckNull(rst.Fields!ToolName)
        
        ' Test to see if the tool number and offset number should be
        ' obtained from the database or the tool file
        If Not clsOptions.UseToolData Then
          .ToolNumber = gvar_CheckNull(rst.Fields!ToolNumber)
          .OffsetNumber = gvar_CheckNull(rst.Fields!ToolOffset)
          .SpindleSpeed = gvar_CheckNull(rst.Fields!SpindleSpeed)
          .CutFeed = gvar_CheckNull(rst.Fields!CutFeed)
          .DownFeed = gvar_CheckNull(rst.Fields!DownFeed)
        End If
        
        .ToolSide = gvar_CheckNull(rst.Fields!ToolSide)
        .ToolSidePartialReverse = gvar_CheckNull(rst.Fields!ToolSidePartialReverse)
        .WidthOfCut = gvar_CheckNull(rst.Fields!WidthOfCut)
        .XYCorners = gvar_CheckNull(rst.Fields!XYCorners)
        .InsertFilePath = gvar_CheckNull(rst.Fields!InsertFilePath)
        .InsertFilePointX = gvar_CheckNull(rst.Fields!InsertFilePointX)
        .InsertFilePointY = gvar_CheckNull(rst.Fields!InsertFilePointY)
        .InsertFileReferencePoint = gvar_CheckNull(rst.Fields!InsertFileReferencePoint)
        .InsertFileGroupNumber = gvar_CheckNull(rst.Fields!InsertParametricGroupNumber)
        
        If IsNull(rst.Fields!CreationMethod) Then
          ' No creation method specified - so use the default - "Manual"
          .CreationMethod = DEF_CREATION_METHOD_MANUAL
        Else
          .CreationMethod = gvar_CheckNull(rst.Fields!CreationMethod)
          If .CreationMethod = "" Then .CreationMethod = DEF_CREATION_METHOD_MANUAL
        End If
        
        .MachiningStyle = gvar_CheckNull(rst.Fields!MachiningStyle)
        
        If IsNull(rst.Fields!CutType) Then
          ' No cut type specified - so use the default "Full"
          .CutType = DEF_CUT_TYPE_FULL
        Else
          ' Test to see if cut type is "" - this will be a path defined in an older
          ' release of AlphaDOOR - so use the default "Full"
          If rst.Fields!CutType = "" Then
            .CutType = DEF_CUT_TYPE_FULL
          Else
            .CutType = gvar_CheckNull(rst.Fields!CutType)
          End If
        End If
        
        .PartialStartElemIndex = gvar_CheckNull(rst.Fields!PartialStartElemIndex)
        .PartialStartElemDist = gvar_CheckNull(rst.Fields!PartialStartElemDist)
        .PartialEndElemIndex = gvar_CheckNull(rst.Fields!PartialEndElemIndex)
        .PartialEndElemDist = gvar_CheckNull(rst.Fields!PartialEndElemDist)
        
        .SlowDownForCorners = gvar_CheckNull(rst.Fields!SlowDownForCorners)
        .DecelerationDistance = gvar_CheckNull(rst.Fields!DecelerationDistance)
        .NumberOfSteps = gvar_CheckNull(rst.Fields!NumberOfSteps)
        .SlowDownTo = gvar_CheckNull(rst.Fields!SlowDownTo)
        .DoNotSlowDownRadius = gvar_CheckNull(rst.Fields!DoNotSlowDownRadius)
        .IgnoreAngleGreaterThan = gvar_CheckNull(rst.Fields!IgnoreAngleGreaterThan)
        .AccelerateOutOfCorner = gvar_CheckNull(rst.Fields!AccelerateOutOfCorner)
        
        .SimpleEngraveFeed = gvar_CheckNull(rst.Fields!SimpleEngraveFeed)
        .SimpleEngraveClearance = gvar_CheckNull(rst.Fields!SimpleEngraveClearance)
        
    End With
    
Exit Function

mbln_FillPathData_Error:
    
    MsgBox Err.Description, vbExclamation, "mbln_FillPathData"
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_FillPathData"
    mbln_FillPathData = False

End Function

Private Sub m_SetPostVariables(dPanel_X As Double, dPanel_Y As Double, dPanel_Z As Double, lProgNum As Long, bPrompt As Boolean)

On Error Resume Next
    
    With App
        
        If Not bPrompt Then
            .SetPostUserVariable "PROGNUM", Format$(CStr(lProgNum), "0000")
            .SetPostUserVariable "DESCRIPTION", Format$(Now, "Medium Date")
        End If
    
        '..set the post variables
        .SetPostUserVariable "ALPHADOOR", "1"
        .SetPostUserVariable "ACAM_CDM", "1"
        .SetPostUserVariable "PANEL_X", CStr(dPanel_X)
        .SetPostUserVariable "PANEL_Y", CStr(dPanel_Y)
        .SetPostUserVariable "PANEL_Z", CStr(dPanel_Z)
        .SetPostUserVariable "ZDIM", CStr(dPanel_Z)

    End With
    
    If (Err.Number <> 0) Then Err.Clear                                               '..04.08.02 - rg
    
End Sub
'EditMark
Private Function mbln_MakeMachining(Door As CDoor, lngGeoNumber As Long, lngPartNumber As Long, _
                                    colANC As Collection, bPreview As Boolean)
                                                                                                               
    Dim rstPath                 As ADODB.Recordset
    Dim pthsToCut               As Paths
    Dim pthsLeads               As Paths
    Dim Pth                     As Path
    Dim pthPath                 As Path
    Dim pthPick                 As Path
    Dim pthBoundary             As Path
    Dim blnHasBoundary          As Boolean
    'Dim sCTX                    As String
    
On Error GoTo mbln_MakeMachining_Error

    '..let's assume everything will be just fine
    mbln_MakeMachining = True
    
    Set rstPath = grst_GetPaths(Door.TypeName)
    
    strCTX = clsOptions.CTXFile
    
    If Not (rstPath Is Nothing) Then
                    
        '..move to the first recordset
        rstPath.MoveFirst
            
        Do Until rstPath.EOF
                            
            '..fill up the path data
            If mbln_FillPathData(rstPath) Then
                
                '..initialize boundary
                blnHasBoundary = False
                
                '..let's see what we're supposed to do with it
                With clsPathData
                    
                    '..is this an inserted toolpath?
                    If .MachineMethod = DEF_MACHINE_METHOD_INSERT Then
                        
                        If Not mbln_InsertARD(Door, bPreview) Then        ' dWidth, dLength) Then         '..07.03.02 - rg
                            
                            '..let the user know and see if they want to continue
                            If MsgBox(Frame.ReadTextFile(strCTX, 500, 22) & Space(1) & UCase$(.InsertFilePath) & Space(3) & _
                                             vbCrLf & Frame.ReadTextFile(strCTX, 500, 20) & Space(3), _
                                             vbQuestion + vbYesNo, Frame.ReadTextFile(strCTX, 120, 1)) = vbNo Then GoTo mbln_MakeMachining_Error
                            
                        End If
                                                                        
                    Else
                    
                        '..get the paths to be cut
                        If (.GroupID <> 0) Then                                     '..04.11.02 - rg
                            
                            '..paths belong in a group so lets group 'em
                            Set pthsToCut = mpths_PathsInGroup(ActiveDrawing.Geometries, .GroupID)
                            
                        Else
                             
                             ' LJO 08.02.11 - modified with fix for outer geometry
                             Set pthsToCut = mpths_PathsNotInGroup(ActiveDrawing.Geometries, .PathOffsetFrom)
                            
                            ' LJO 26.02.03 - to fix problem with ignore outer geometry and
                            ' outer geometry toolpaths exist in database.
                            
                            If Not Door.IgnoreOuterGeometry Then
                              '..paths are not grouped or are from old type
                              Set pthsToCut = mpths_PathsNotInGroup(ActiveDrawing.Geometries, .PathOffsetFrom)
                            Else
                              Set pthsToCut = Nothing
                            End If
                            
                        End If
                                                                
                        
                        If Not pthsToCut Is Nothing Then
                        
                          '..let's go thru all the paths we're supposed to deal with
                          For Each pthPath In pthsToCut
                          
                              '..are we offsetting?
                              If (.PathOffsetValue <> 0) Then
                                      
                                  '..up the geo count by 1
                                  lngGeoNumber = (lngGeoNumber + 1)
                                      
                                  '..prep machining for this path
                                  If Not mbln_MachineWithOffset(clsPathData, pthPath, _
                                                                pthPick, pthBoundary, blnHasBoundary) Then
                                      mbln_MakeMachining = False
                                      GoTo Controlled_Exit
                                  End If
      
                              Else
                                  
                                  '..test for partial cut
                                  If .CutType = DEF_CUT_TYPE_PARTIAL Then
                                      
                                      Set pthPick = DrawParametricPartialPath(pthPath, .PartialStartElemIndex, .PartialStartElemDist, .PartialEndElemIndex, .PartialEndElemDist, .ToolSidePartialReverse)
                                  
                                  Else
                                      
                                      '..assign the active geo
                                      Set pthPick = pthPath
                                      
                                  End If
                                      
                                  '..not offsetting this path
                                  If Not mbln_MachineWithNoOffset(clsPathData, pthPick, _
                                                                  pthBoundary, blnHasBoundary) Then
                                      mbln_MakeMachining = False
                                      GoTo Controlled_Exit
                                  End If
      
                                  
                              
                              
                              End If
                                                                      
                              '..we've got the path we need to let's set the parameters
                                                                      
                              '..reverse open geo?
                              If .ToolDirectionIsReversed Then pthPick.Reverse
                                                                          
                              '..check for closed geo
                              If pthPick.Closed = True Then
                                  
                                  '..set to inside/outside
                                  pthPick.ToolInOut = .ToolInOut
                              
                              Else
                                  
                                  '..set to left/right
                                  pthPick.ToolSide = .ToolSide
                              
                              End If
                          
                              '..assign tool direction
                              If .LeadEntryPointIsCorner <> AdoorLeadEntryPoint_Drawn Then
                                pthPick.CW = .ToolDirectionCW
                              End If
  
                              '..select it
                              pthPick.Selected = True
                                          
                              '..setup the Boundary toolside and dir
                              If blnHasBoundary Then
                                  
                                  pthBoundary.CW = Not .ToolDirectionCW
                                  
                                  '..select it
                                  pthBoundary.Selected = True
                              
                              End If
                                                                              
                              '..let's see what we got
                              'With ActiveDrawing
                              '    .ZoomAll
                              '    .Redraw
                              'End With
                              
                              '..see if we need to get the tool (not applicable to machining styles)
                              If .CreationMethod = DEF_CREATION_METHOD_MANUAL Then
                                '..get the tool
                                  If Not gbln_GetTool(.ToolFullPath) Then
                                      
                                      MsgBox Frame.ReadTextFile(strCTX, 500, 24) & Space(3), vbExclamation, Frame.ReadTextFile(strCTX, 120, 1)
                                      mbln_MakeMachining = False
                                      GoTo Controlled_Exit
                                      
                                  End If
                              End If
                          
                              '..machine it
                              If gbln_MakeMillData(pthsLeads, False, bPreview) Then
                                              
                                  '..if rough/finish the apply leads
                                  If (.MachineMethod = DEF_MACHINE_METHOD_ROUGHFINISH _
                                    Or .MachineMethod = DEF_MACHINE_METHOD_ENGRAVE _
                                    Or .MachineMethod = DEF_MACHINE_METHOD_POCKET _
                                    Or .MachineMethod = DEF_MACHINE_METHOD_SIMPLE_ENGRAVE) Then
                                  
                                      '..apply leads
                                      If Not gbln_ApplyLeads(pthsLeads, pthPick.Closed) Then
                                          
                                          MsgBox Frame.ReadTextFile(strCTX, 500, 25) & Space(1) & .OperationNumber & Space(3), _
                                                  vbExclamation, "gbln_MakeMachining"
                                                  
                                          mbln_MakeMachining = False
                                          GoTo Controlled_Exit
                                                         
                                      End If
                                  
                                  End If
                                  
                                  '..apply attributes
                                  If Not bPreview Then
                                  If App.ActiveDrawing.ToolPaths.Count > 0 Then
                                   For Each Pth In App.ActiveDrawing.ToolPaths
                                   Call m_SetAttributes(Pth, Door)
                                   Next
                                   End If

                                    Call m_SetAttributes(pthsLeads(1), Door)    ' sType, uDTD.Width, uDTD.Length)         ' dWidth, dLength) Then         '..07.09.02 - rg
                                  End If
                                   
                                  
                                  
                                  
                                  
                                  '..test for corner slow down
                                  If .MachineMethod = DEF_MACHINE_METHOD_ROUGHFINISH And .SlowDownForCorners Then
                                      
                                      For Each Pth In pthsLeads
                                          Pth.SlowDownForCorners .DecelerationDistance, .NumberOfSteps, .SlowDownTo, .DoNotSlowDownRadius, .IgnoreAngleGreaterThan, .AccelerateOutOfCorner
                                      Next
                                      
                                  End If
                                  
                                  '..let's see the damn thing
                                  'ActiveDrawing.ZoomAll
                                  
                                  '..deselect the picked geo
                                  pthPick.Selected = False
                                  
                                  '..deselect the boundary geo
                                  If blnHasBoundary Then pthBoundary.Selected = False
                                          
                              Else
                                  
                                  '..woops
                                  MsgBox Frame.ReadTextFile(strCTX, 500, 26) & Space(1) & .OperationNumber & Space(3), _
                                         vbExclamation, "gbln_MakeMachining"
      
                                  '..better stop
                                  mbln_MakeMachining = False
                                  
                                  '..deselect the picked geo
                                  pthPick.Selected = False
                                  
                                  '..deselect the boundary geo
                                  If blnHasBoundary Then pthBoundary.Selected = False
                                  
                                  GoTo Controlled_Exit
                                  
                              End If
                                      
                              If pthPick.Attribute(DEF_ATT_TEMP_PATH) = "1" Then
                                pthPick.Selected = True
                                ActiveDrawing.DeleteSelected
                              End If
                                      
                          Next pthPath
                                    
                        End If
                        
                    End If
                                    
                End With
                    
            End If
              
            '..move to the next recordset
            rstPath.MoveNext
                                
        Loop
        
    End If
    
    '..add handle drilling
    If Not mbln_DrawHandleHoles(Door, bPreview) Then
    
        mbln_MakeMachining = False
        GoTo Controlled_Exit
    
    End If
    
    ' TFS#62973
    ' Exclude generic cutom macro for Chinacam
    ' (Chinacam does not allow add-ins to be enabled programatically,
    ' apart from CDM User Styles)
    If App.ProgramLevel <> acamLevelCHINACAM Then
    
        ' TFS #56993 - Run a custom macro
        If clsOptions.CustomMacro <> "" Then
          If Not mbln_RunCustomMacro(clsOptions.CustomMacro, Door) Then
            mbln_MakeMachining = False
            GoTo Controlled_Exit
          End If
        End If
        
    End If
    
    
    '..update the database and save the file
    mbln_MakeMachining = mbln_UpdateAndSave(Door, lngPartNumber, colANC, bPreview)
        
Controlled_Exit:
    
    Set pthPath = Nothing
    Set pthPick = Nothing
    Set pthsLeads = Nothing
    Set pthsToCut = Nothing

Exit Function

mbln_MakeMachining_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_MakeMachining"
    mbln_MakeMachining = False
    Resume Controlled_Exit

End Function

Private Function mbln_UpdateAndSave(Door As CDoor, lPartNumber As Long, _
                                   colANC As Collection, bPreview As Boolean) As Boolean
    
    Dim strSave                 As String
    Dim strFilename             As String
    Dim P                       As Path
    Dim ps                      As Paths
    Dim strImage                As String
    Dim pGeos                   As Paths
    Dim dblPartRecoveryX        As Double
    Dim dblPartRecoveryY        As Double
    Dim blnPartRecovery         As Boolean
    Dim blnPartRecoveryNoRotate As Boolean
    Dim blnRotateDoor           As Boolean
    Dim blnTranslateDoor        As Boolean
    
On Error GoTo mbln_UpdateAndSave_Error
        
    '..get the toolpaths
    Set ps = ActiveDrawing.ToolPaths
              
    '..start out ok
    mbln_UpdateAndSave = True
    
    '..if doing a preview then get out now
    If bPreview Then GoTo Controlled_Exit

    SuppressUpdateRapids False

    dblPartRecoveryX = clsOptions.PartRecoveryOffsetX
    dblPartRecoveryY = clsOptions.PartRecoveryOffsetY
    blnPartRecoveryNoRotate = clsOptions.PartRecoveryIgnoreGrain
    
    blnPartRecovery = (dblPartRecoveryX <> 0) Or (dblPartRecoveryY <> 0)
    
    '..apply any router rules that are applicable
    '..only apply for Alphadoor user styles
    If Door.StyleNumber = 930 Then
      If Not mbln_RulesApply(Door) Then
        mbln_UpdateAndSave = False
        GoTo Controlled_Exit
      End If
    ElseIf Door.StyleNumber = 900 Then
      ' Check for rotation restriction from the drawing file
      If ActiveDrawing.Attribute(DEF_ATT_NEST_RESTRICT) <> "" Then
        Door.RotationMethod = ActiveDrawing.Attribute(DEF_ATT_NEST_RESTRICT)
      End If
    End If
    
    '..check for grain and flip if needed
    With clsNest

        '..check rotation method and assign proper angle                            '..07.03.02 - rg
        Select Case Door.RotationMethod
        
            Case adoorPART_ROTATION_FREE
                
                '..should have already assigned angle, so leave as is
                
                ' Test for part recovery options
                If blnPartRecovery Then
                  g_PartRecovery dblPartRecoveryX, dblPartRecoveryY
                End If
            
            Case adoorPART_ROTATION_LOCKX
            
                If blnPartRecovery And blnPartRecoveryNoRotate Then
                    blnRotateDoor = True
                Else
                  '..rotate part
                  Set pGeos = ActiveDrawing.Geometries
  
                  For Each P In ps
                      P.RotateL 90, 0, 0
                  Next P
  
                  For Each P In pGeos
                      P.RotateL 90, 0, 0
                  Next
                End If
                                
                '..lock rotation at 0
                Door.RotationAngle = 0
                
                If blnPartRecovery Then
                  g_PartRecovery dblPartRecoveryX, dblPartRecoveryY
                End If
                
                ' Set a flag if using bypass nest option to indicate door has been rotated
                ' and will need translation to +ve x and +ve y quadrant
                If Door.ByPassNest Then
                    blnTranslateDoor = True
                End If
                
                ActiveDrawing.AutoScale ' .ZoomAll
                
            Case adoorPART_ROTATION_LOCKY
                
                ' Test for part recovery options
                If blnPartRecovery Then
                  g_PartRecovery dblPartRecoveryX, dblPartRecoveryY
                End If
                
                '..already orientated in Y so just lock the angle
                Door.RotationAngle = 0
                
            Case adoorPART_ROTATION_MATERIAL
                                    
                If Not .AllowRotation Then
                
                    If .GrainInX Then
                        
                        If blnPartRecovery And blnPartRecoveryNoRotate Then
                            blnRotateDoor = True
                        Else
                          '..rotate part
                          Set pGeos = ActiveDrawing.Geometries
          
                          For Each P In ps
                              P.RotateL 90, 0, 0
                          Next P
          
                          For Each P In pGeos
                              P.RotateL 90, 0, 0
                          Next
                        End If
                                        
                        '..lock rotation at 0
                        Door.RotationAngle = 0

                        ' Set a flag if using bypass nest option to indicate door has been rotated
                        ' and will need translation to +ve x and +ve y quadrant
                        If Door.ByPassNest Then
                            blnTranslateDoor = True
                        End If

                        ActiveDrawing.AutoScale ' .ZoomAll
                        
                    Else
                    
                        ' Test for part recovery options
                        If blnPartRecovery Then
                          g_PartRecovery dblPartRecoveryX, dblPartRecoveryY
                        End If
                    
                    End If
                    
                    '..lock rotation at 0                                           '..07.03.02 - rg
                    Door.RotationAngle = 0
                    
                Else
                
                    ' Test for part recovery options
                    If blnPartRecovery Then
                      g_PartRecovery dblPartRecoveryX, dblPartRecoveryY
                    End If
                    
                    Door.RotationAngle = 90
                  
                End If
        
            End Select

    End With

    With Door
        
        '..assign the name  ~~ always add one because should always be one behind ~~
        strFilename = gstr_JobName & DEF_UNDERSCORE & clsTypeData.TypeName & DEF_UNDERSCORE & (lPartNumber + 1)
        
        If clsOptions.OutputToSecondLocation Then
        
            If clsOptions.OutputSecondCreateSubFolder Then
              strSave = gstr_EnsureBackslash(clsOptions.PathToRootSecond) & gstr_JobName & "\" & strFilename
            Else
              strSave = gstr_EnsureBackslash(clsOptions.PathToRootSecond) & strFilename
            End If
            
            If clsOptions.OutputSecondDoorDrawings Then
                ActiveDrawing.SaveAs strSave & DEF_EXTENSION_ARD
            End If
        
        End If
        
        '..assign the full path to ard/anc file
        If clsOptions.OutputResultsSubFolder Then
          strSave = gstr_EnsureBackslash(clsOptions.PathToRoot) & gstr_JobName & "\" & strFilename
        Else
          strSave = gstr_EnsureBackslash(clsOptions.PathToRoot) & strFilename
        End If
        
        '..assign the full path to the image file
        strImage = gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE
        
        If Right$(strImage, 1) <> DEF_BACKSLASH Then strImage = strImage & DEF_BACKSLASH
        
        strImage = strImage & strFilename
                                                                
        '..check for nesting
        If Not .ByPassNest Then
                                                                
            '..loop thru all toolpaths and assign the anc path to an attribute (used for labeling)
            For Each P In ps

                With P
                    .Attribute(DEF_ATT_ANC_NAME) = strFilename & "." & clsOptions.NCFileExtension
                    .Attribute(DEF_ATT_ANC_FULLNAME) = strSave & "." & clsOptions.NCFileExtension
                    .Attribute(DEF_ATT_PART_IMAGE) = strImage
                End With

            Next P
            
            '..output nc code for single door?
            If clsOptions.SaveAllDoorNC Then
                                                                                                      
                '..set the post variables
                Call m_SetPostVariables(.Width, .Length, clsNest.SheetThickness, _
                                        lPartNumber + 1, clsOptions.PromptForProgram)
                                                                                                                                                                                                                                                          
                If clsOptions.OutputMPR Then
                    m_MprSave CStr(.Length), CStr(.Width), CStr(clsNest.SheetThickness), "0", "0", "0", "0", "0", "0", "0", "", "0", strSave & DEF_EXTENSION_MPR
                Else
                  ActiveDrawing.OutputNC strSave & "." & clsOptions.NCFileExtension, acamOutNcFILE, False
                
                  '..add single door anc file to array
                  colANC.add strSave & "." & clsOptions.NCFileExtension
                End If
                
            End If
                
            '..save it
            If clsOptions.OutputToSecondLocation Then
                
                If clsOptions.OutputSecondCreateSubFolder Then
                  strSave = gstr_EnsureBackslash(clsOptions.PathToRootSecond) & gstr_JobName & "\" & strFilename
                Else
                  strSave = gstr_EnsureBackslash(clsOptions.PathToRootSecond) & strFilename
                End If
                
                If clsOptions.OutputSecondDoorNCFiles Then
                                
                    If clsOptions.OutputMPR Then
                        m_MprSave CStr(.Length), CStr(.Width), CStr(clsNest.SheetThickness), "0", "0", "0", "0", "0", "0", "0", "", "0", strSave & DEF_EXTENSION_MPR
                    Else
                      ActiveDrawing.OutputNC strSave & "." & clsOptions.NCFileExtension, acamOutNcFILE, False
                    
                      '..add single door anc file to array
                      colANC.add strSave & clsOptions.NCFileExtension
                    End If
                End If
            
            End If
                
            ' Assign the allowable door rotation angle to the door toolpaths
            ' This is used mainly for twin head nesting
            For Each P In ps
                P.Attribute(DEF_ATT_DOOR_ROTATION_ANGLE) = .RotationAngle
            Next

            
        Else    '..always output nc file when doing single parts
                
            '..set the post variables
            Call m_SetPostVariables(.Width, .Length, clsNest.SheetThickness, _
                                    lPartNumber + 1, clsOptions.PromptForProgram)
                                                                                                       
            '..insert the path to anc
            '.Fields!PathToANC = strFilename & DEF_EXTENSION_ANC 'strSave
                
            '..loop thru all toolpaths and assign the anc path to an attribute (used for labeling)
            For Each P In ps

                With P
                    .Attribute(DEF_ATT_ANC_NAME) = strFilename & clsOptions.NCFileExtension
                    .Attribute(DEF_ATT_PART_IMAGE) = strImage
                End With

            Next P
            
            If blnTranslateDoor Then
            
                '..rotate part
                Set pGeos = ActiveDrawing.Geometries
        
                For Each P In ps
                    P.MoveL Door.Length, 0
                Next P
        
                For Each P In pGeos
                    P.MoveL Door.Length, 0
                Next
            
            End If
            
            '..save it
            If clsOptions.OutputToSecondLocation Then
                
                If clsOptions.OutputSecondCreateSubFolder Then
                  strSave = gstr_EnsureBackslash(clsOptions.PathToRootSecond) & gstr_JobName & "\" & strFilename
                Else
                  strSave = gstr_EnsureBackslash(clsOptions.PathToRootSecond) & strFilename
                End If
                
                If clsOptions.OutputMPR Then
                    m_MprSave CStr(.Length), CStr(.Width), CStr(clsNest.SheetThickness), "0", "0", "0", "0", "0", "0", "0", "", "0", strSave & DEF_EXTENSION_MPR
                Else
                    strFilename = gstr_JobName & DEF_UNDERSCORE & clsTypeData.TypeName & DEF_UNDERSCORE & .Width & "x" & .Length & DEF_UNDERSCORE & (lPartNumber + 1)
                    ActiveDrawing.OutputNC strSave & "." & clsOptions.NCFileExtension, acamOutNcFILE, False
                End If
            
            End If
            
            If clsOptions.OutputMPR Then
                m_MprSave CStr(.Length), CStr(.Width), CStr(clsNest.SheetThickness), "0", "0", "0", "0", "0", "0", "0", "", "0", strSave & DEF_EXTENSION_MPR
            Else
              strFilename = gstr_JobName & DEF_UNDERSCORE & clsTypeData.TypeName & DEF_UNDERSCORE & .Width & "x" & .Length & DEF_UNDERSCORE & (lPartNumber + 1)
              
              '..assign the full path to ard/anc file
              If clsOptions.OutputResultsSubFolder Then
                strSave = gstr_EnsureBackslash(clsOptions.PathToRoot) & gstr_JobName & "\" & strFilename
              Else
                strSave = gstr_EnsureBackslash(clsOptions.PathToRoot) & strFilename
              End If
              
              ActiveDrawing.OutputNC strSave & "." & clsOptions.NCFileExtension, acamOutNcFILE, False
              
              '..add single door anc file to array
              colANC.add strSave & "." & clsOptions.NCFileExtension
            End If
            
        End If
         
                  
        If Not clsOptions.DisableReports Then
                
            '..output the emf for the reports
            ActiveDrawing.SaveEmfFile strImage & DEF_EXTENSION_EMF, False, False

        End If
        
                
        If clsOptions.OutputResultsSubFolder Then
          strSave = gstr_EnsureBackslash(clsOptions.PathToRoot) & gstr_JobName & "\" & strFilename
        Else
          strSave = gstr_EnsureBackslash(clsOptions.PathToRoot) & strFilename
        End If
                
                
        ' Test for rotation
        If blnRotateDoor Then
        
            '..rotate part
            Set pGeos = ActiveDrawing.Geometries
    
            For Each P In ps
                P.RotateL 90, 0, 0
            Next P
    
            For Each P In pGeos
                P.RotateL 90, 0, 0
            Next
        
        Else
        
          If blnPartRecovery Then
            ' Move the door into the positive x and positive y quadrant
            ' Because this door has been rotated, it will need to be shifted
            ' in the x-axis by the length
            g_PartRecovery dblPartRecoveryX, dblPartRecoveryY
          
          End If
        
        End If
        
        '..save it to default location
        ActiveDrawing.SaveAs strSave & DEF_EXTENSION_ARD
        
        '..update the database
''        .Fields!PathToARD = strFilename & DEF_EXTENSION_ARD 'strSave
        
        '..update it
        '**********************************
        '.Update
                                   
                                   
        Dim NestZone As CNestingZone
        Dim dblPartArea As Double
                        
        ' TFS#79915
        If clsNest.NestingOption = adoorNESTING_RADNEST Then
                        
            ' Test to see if the component has been assigned a zone
            If Door.NestingZone <> 0 Then
              ' Manual zoning for this component has been specified - do nothing
            Else
              ' Test to see if any automatic zoning conditions have been defined
              dblPartArea = Door.Width * Door.Length
              For Each NestZone In colAutoNestZones
                If (Door.Width <= NestZone.ZonePartDimension) Or (Door.Length <= NestZone.ZonePartDimension) Or (dblPartArea <= NestZone.ZonePartArea) Then
                  Door.NestingZone = NestZone.ZoneNumber
                  Door.NestingZoneID = NestZone.ZoneID
                  NestZone.AutoZoneInUse = True
                End If
              Next
            End If
        End If
                                   
                                   
        '..add it to the nest list
        If Not .ByPassNest Then Call clsNest.AddPart(strSave, .Quantity, .RotationAngle, .NestingPriority, .NestingZone)

        ' If we are not saving individual door drawings then add the filename
        ' to a collection and delete after nesting
        If Not clsOptions.SaveAllDoorARD Then
          colDeleteFiles.add strSave & DEF_EXTENSION_ARD
        End If

        '..increment part count here, it is done here in case none get completed
        lPartNumber = lPartNumber + 1
    
    End With
    
Controlled_Exit:

    Set ps = Nothing
    
Exit Function

mbln_UpdateAndSave_Error:
    
    With Door
    
        '..tell the user and ask to continue
        MsgBox Frame.ReadTextFile(strCTX, 500, 18) & Space(1) & .StyleNumber & _
               DEF_COMMA & Frame.ReadTextFile(strCTX, 500, 19) & Space(1) & .TypeName & _
               vbCrLf & Frame.ReadTextFile(strCTX, 500, 20), _
               vbExclamation + vbYesNo, Frame.ReadTextFile(strCTX, 120, 1)
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_UpdateAndSave"
    mbln_UpdateAndSave = False
    Resume Controlled_Exit
    
End Function

Private Function mbln_UpdateAndSavePress(Door As CDoor, lPartNumber As Long) As Boolean
    
    Dim strSave                 As String
    Dim strFilename             As String
    Dim P                       As Path
    Dim strImage                As String
    Dim pGeos                   As Paths
    
On Error GoTo mbln_UpdateAndSave_Error
        
    '..get the toolpaths
    Set pGeos = ActiveDrawing.Geometries
                
    '..start out ok
    mbln_UpdateAndSavePress = True

    '..check for grain and flip if needed
    With clsNest

        '..check rotation method and assign proper angle                            '..07.03.02 - rg
        Select Case Door.RotationMethodPress
        
            ' Values will be
            '
            ' 0=Not set (no rotation restriction)
            ' 1=Lock X
            ' 2=Lock Y
                
            Case adoorPART_ROTATION_LOCKX
                                            
                For Each P In pGeos
                    P.RotateL 90, 0, 0
                Next
                                
                ActiveDrawing.ZoomAll
                
                '..lock rotation at 0
                Door.RotationAngle = 0
                
            Case adoorPART_ROTATION_LOCKY
                
                '..already orientated in Y so just lock the angle
                Door.RotationAngle = 0
                
            Case Else
                                    
                ' No press rotation restrictions
                ' Test to see if there is a colour rotation restriction
                
                Select Case Door.ColourRotationMethod
                
                    Case adoorPART_ROTATION_MATERIAL
                    
                        Select Case Door.ColourDefinedRotationMethod
                        
                            Case adoorPART_ROTATION_LOCKX
                                                            
                                For Each P In pGeos
                                    P.RotateL 90, 0, 0
                                Next
                        
                                ' Now lock the angle
                                Door.RotationAngle = 0
                        
                            Case adoorPART_ROTATION_LOCKY
                                
                                '..already orientated in Y so just lock the angle
                                Door.RotationAngle = 0
                        
                            Case Else
                            
                                ' Free rotation
                                Door.RotationAngle = 90
                        
                        End Select
                    
                    
                    Case adoorPART_ROTATION_LOCKX
                                                    
                        For Each P In pGeos
                            P.RotateL 90, 0, 0
                        Next
                
                        ' Now lock the angle
                        Door.RotationAngle = 0
                    
                    Case adoorPART_ROTATION_LOCKY
                        
                        '..already orientated in Y so just lock the angle
                        Door.RotationAngle = 0
                
                    Case adoorPART_ROTATION_FREE
                    
                        ' Free rotation
                        Door.RotationAngle = 90
                
                End Select
        
            End Select

    End With

    With Door

        If .StyleNumber = 900 Then
          
          '..assign the name  ~~ always add one because should always be one behind ~~
          strFilename = .JobName & DEF_UNDERSCORE & _
                        .PressName & DEF_UNDERSCORE & _
                        .FoilColour & DEF_UNDERSCORE & _
                        .DoorThickness & DEF_UNDERSCORE & _
                        gstr_ParseName(.TypeName) & DEF_UNDERSCORE & (lPartNumber + 1)
        
        ElseIf .StyleNumber = 930 Then
          
          '..assign the name  ~~ always add one because should always be one behind ~~
          strFilename = .JobName & DEF_UNDERSCORE & _
                        .PressName & DEF_UNDERSCORE & _
                        .FoilColour & DEF_UNDERSCORE & _
                        .DoorThickness & DEF_UNDERSCORE & _
                        .TypeName & DEF_UNDERSCORE & (lPartNumber + 1)
        
        End If
    End With
        
    '..assign the full path to ard/anc file
    If clsOptions.OutputResultsSubFolder Then
      strSave = gstr_EnsureBackslash(clsOptions.PathToRoot) & gstr_JobName & "\" & strFilename
    Else
      strSave = gstr_EnsureBackslash(clsOptions.PathToRoot) & strFilename
    End If
    
    '..assign the full path to the image file
    strImage = gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE
    
    If Right$(strImage, 1) <> DEF_BACKSLASH Then strImage = strImage & DEF_BACKSLASH
    
    strImage = strImage & strFilename
                                                                                                                     
    '..loop thru all toolpaths and assign the anc path to an attribute (used for labeling)
    For Each P In pGeos

        With P
            .Attribute(DEF_ATT_PART_IMAGE) = strImage
        End With

    Next P
                    
    '..save it to default location
    ActiveDrawing.SaveAs strSave & DEF_EXTENSION_ARD
        
    '..add it to the nest list
    With Door
      clsNest.AddPart strSave, .Quantity, .RotationAngle, .NestingPriority, 0
    End With

    If Not clsOptions.SaveAllDoorARD Then
      colDeleteFiles.add strSave & DEF_EXTENSION_ARD
    End If
  
    '..increment part count here, it is done here in case none get completed
    lPartNumber = lPartNumber + 1
    
Controlled_Exit:

    Set pGeos = Nothing
    
Exit Function

mbln_UpdateAndSave_Error:
    
    With Door
    
        '..tell the user and ask to continue
        MsgBox Frame.ReadTextFile(strCTX, 500, 18) & Space(1) & .StyleNumber & _
               DEF_COMMA & Frame.ReadTextFile(strCTX, 500, 19) & Space(1) & .TypeName & _
               vbCrLf & Frame.ReadTextFile(strCTX, 500, 20), _
               vbExclamation + vbYesNo, Frame.ReadTextFile(strCTX, 120, 1)
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_UpdateAndSavePress"
    mbln_UpdateAndSavePress = False
    Resume Controlled_Exit
    
End Function


Public Sub m_SmallFirst()

    '' Rectangular nesting (rectangular nesting does not place small components last)
    '' This will re-order all external toolpaths, and cut the smallest area components first
    '' Spacemaker & OS Doors were finding that small components were moving due to loss of vacuum
    '' at the end of a nested sheet

    SuppressUpdateRapids True

    OrderExternalToolpaths

    SuppressUpdateRapids False
    
End Sub

Private Function mbln_InsertFillerParts(Material As CMaterial) As Boolean

    Dim nstInfo             As NestInformation
    Dim nstFillerSheet      As NestSheet
    Dim nData               As NestData
    Dim pSheet              As Path
    Dim FSO                 As Scripting.FileSystemObject
    
On Error GoTo mbln_InsertFillerParts_Error
    
    '..so far, so good
    mbln_InsertFillerParts = True
                                      
    Set FSO = New Scripting.FileSystemObject
        
    '..anything there?
    If Not FSO.FileExists(LicomdirPath & Material.InsertFillerFile) Then
        '..no good
        Material.InsertFillerFile = ""
    Else
        Material.InsertFillerFile = gstr_StripExtension(LicomdirPath & Material.InsertFillerFile)
    End If
      
    If Not FSO.FileExists(LicomdirPath & Material.InsertFillerFile2) Then
        '..no good
        Material.InsertFillerFile2 = ""
    Else
        Material.InsertFillerFile2 = gstr_StripExtension(LicomdirPath & Material.InsertFillerFile2)
    End If
      
    If Not FSO.FileExists(LicomdirPath & Material.InsertFillerFile3) Then
        '..no good
        Material.InsertFillerFile3 = ""
    Else
        Material.InsertFillerFile3 = gstr_StripExtension(LicomdirPath & Material.InsertFillerFile3)
    End If
    
    If Material.InsertFillerFile = "" And Material.InsertFillerFile2 = "" And Material.InsertFillerFile3 = "" Then
      MsgBox Frame.ReadTextFile(strCTX, 500, 56), vbExclamation, DEF_PROJECT_NAME
      mbln_InsertFillerParts = False
      GoTo Controlled_Exit
    End If
                       
    Set nstInfo = ActiveDrawing.GetNestInformation

    '..loop thru each sheet filling in where it can
    For Each nstFillerSheet In nstInfo.Sheets
        
        Set pSheet = nstFillerSheet.Geometry
        
        With clsNest
        
            .StartNestListRouter Material.MaterialName & DEF_UNDERSCORE & "FILL", Material, True, False
            
            If Material.InsertFillerFile <> "" Then
              .AddPart Material.InsertFillerFile, 500, 90, 1, 0   ' -> made a large number to emmulate MAX            '..07.26.02 - rg
            End If
            
            If Material.InsertFillerFile2 <> "" Then
              .AddPart Material.InsertFillerFile2, 500, 90, 1, 0
            End If
            
            If Material.InsertFillerFile3 <> "" Then
              .AddPart Material.InsertFillerFile3, 500, 90, 1, 0
            End If
        
            Set nData = ActiveDrawing.CreateNestData(.NestListName)
            
        End With
        
        With nData
                        
            '..user defined
            .AddSheet pSheet, clsNest.SheetName, clsNest.SheetThickness, 1
            .Direction = nestTOPRIGHT
            .EdgeGap = clsNest.GapAtSheetEdge
            .Gap = clsNest.MinimumGap
            .LeadGap = clsNest.ExtraGapAtLead
            .Resolution = clsNest.SearchResolution
            .Subroutines = clsNest.NCSubroutines
            .MergeTools = CLng(clsNest.MergeTools)
            .OrderInnerFirst = CLng(clsNest.InnerRoutesFirst)
                
            '..nest it
            .DoNest
    
        End With
        
    Next nstFillerSheet
    
    '..save nest list?
    If Not clsOptions.SaveAllNestANL Then
        
        With FSO
            If .FileExists(clsNest.NestListName) Then .DeleteFile (clsNest.NestListName)
        End With
    
    End If
    
Controlled_Exit:
    
    Set FSO = Nothing
    Set nstInfo = Nothing
    Set nData = Nothing

Exit Function

mbln_InsertFillerParts_Error:

    mbln_InsertFillerParts = False
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_InsertFillerParts"
    Resume Controlled_Exit
    


''' *************** Alternative Method For Filler Part Insertion ******************
'''
'''    Dim nstInfo             As NestInformation
'''    Dim nstFillerSheet      As NestSheet
'''    Dim nData               As NestData
'''    Dim strFiller           As String
'''    Dim pSheet              As Path
'''    Dim FSO                 As Scripting.FileSystemObject
'''    Dim ShtList             As Sheetlist
'''    Dim N                   As Nesting
'''    Dim Nl                  As Nestlist
'''
'''On Error GoTo mbln_InsertFillerParts_Error
'''
'''    '..so far, so good
'''    mbln_InsertFillerParts = True
'''
'''    Set FSO = New Scripting.FileSystemObject
'''
'''    '..remove the file extension
'''    strFiller = clsNest.InsertFillerFile
'''
'''    '..anything there?
'''    If Len(Trim$(strFiller)) = 0 Then
'''
'''        '..no good
'''        mbln_InsertFillerParts = False
'''        GoTo Controlled_Exit
'''
'''    '..make sure file is there
'''    ElseIf Not FSO.FileExists(strFiller) Then
'''
'''        '..no good
'''        mbln_InsertFillerParts = False
'''        GoTo Controlled_Exit
'''
'''    End If
'''
'''    '..remove the extension
'''    strFiller = Left$(strFiller, Len(strFiller) - 4)
'''
'''    Set nstInfo = Drw.GetNestInformation
'''
'''    '..loop thru each sheet filling in where it can
'''    For Each nstFillerSheet In nstInfo.Sheets
'''
'''        Set ShtList = N.NewSheetList
'''
'''        ShtList.Add nstFillerSheet.Path
'''
'''    Next
'''
'''
'''    With clsNest
'''
'''        .StartNestList sJob, rMat.Fields!Name & DEF_UNDERSCORE & "FILL", rMat, False
'''        .AddPart strFiller, 500, 90, 1  ' -> made a large number to emmulate MAX            '..07.26.02 - rg
'''
'''        Set Nl = N.LoadNestList(.NestListName)
'''
'''    End With
'''
'''    With Nl
'''
'''        '..user defined
'''        '.AddSheet pSheet, clsNest.SheetName, clsNest.SheetThickness, 1
'''        .NestSide = clsNest.PackTo
'''        .EdgeGap = clsNest.GapAtSheetEdge
'''        .PartGap = clsNest.MinimumGap
'''        .LeadInGap = clsNest.ExtraGapAtLead
'''        .Resolution = clsNest.SearchResolution
'''        .UseSubroutines = clsNest.NCSubroutines
'''        .MinimiseToolChanges = CLng(clsNest.MergeTools)
'''        .InnerFirst = CLng(clsNest.InnerRoutesFirst)
'''
'''        .Save
'''
'''        '..constants !! PRODUCES ERROR !!
'''        '.RepeatFirstRowOrColumn = False
'''
'''        '..nest it
'''        '.DoNest
'''
'''    End With
'''
'''    N.Nest Nl, ShtList
'''
'''    '..save nest list?
'''    If Not clsOptions.SaveAllNestANL Then
'''
'''        With FSO
'''            If .FileExists(clsNest.NestListName) Then .DeleteFile (clsNest.NestListName)
'''        End With
'''
'''    End If
'''
'''Controlled_Exit:
'''
'''    Set FSO = Nothing
'''    Set nstInfo = Nothing
'''    Set nData = Nothing
'''
'''Exit Function
'''
'''mbln_InsertFillerParts_Error:
'''
'''    mbln_InsertFillerParts = False
'''    If (Err.Number <> 0) Then WriteError Err, True, "mbln_InsertFillerParts"
'''    Resume Controlled_Exit
'''
End Function

Private Sub m_InsertReportDataPress(Press As CPress, PressColour As CPressColour, Optional PressThickness As Double)
                                   
    Dim Ni                      As NestPartInstance
    Dim niPart                  As NestPartInstance
    Dim nInfo                   As NestInformation
    Dim SH                      As NestSheet
    Dim psNI                    As Paths
    Dim pNI                     As Path
    Dim strNestImage            As String
    Dim strPartImage            As String
    Dim strType                 As String
    Dim lngItem                 As Long
    Dim lngQuantityOnSheet      As String
    Dim lngDetailID             As Long                                             '..07.09.02 - rg
    Dim lngOrderID              As Long
    Dim sSql                    As String
    Dim lngSheetNumber          As Long
    Dim strPressFileName        As String
    Dim lngCustomerID           As Long
    Dim PressReportSheet        As CPressReportSheet
    Dim PressReportDataItem     As CPressReportDataItem

On Error GoTo m_InsertReportData_Error
                       
    '..ensure connection
    If Not gbln_ConnectToDB Then GoTo Controlled_Exit
        
    If PressThickness <> 0 Then
      strPressFileName = Press.PressName & DEF_UNDERSCORE & PressColour.ColourName & PressThickness
    Else
      strPressFileName = Press.PressName & DEF_UNDERSCORE & PressColour.ColourName
    End If
    
    ' 12 Aug 13 - Moved to end of routine to support Alphacam Reports Engine
    ' (Ensures press nesting attributes are saved)
    'm_CreateAlphaCAMDrawingsOfSheetsPress strPressFileName
        
    '..assign NInfo to current drawing nest information
    Set nInfo = ActiveDrawing.GetNestInformation
    
    '..update the process
    With Frame
      .ShowProgressBox .ReadTextFile(strCTX, 120, 1), .ReadTextFile(strCTX, 300, 22) & Space(3)
    End With
    
    DoEvents
    
    '..get information from each nested sheet
    For Each SH In nInfo.Sheets
        
        lngSheetNumber = lngSheetNumber + 1
        
        '..update the progress
        Frame.SetProgressText Frame.ReadTextFile(strCTX, 300, 22) & Space(3) & SH.Name & Space(3)
        DoEvents

        '..set the path to the wmf file
        strNestImage = gstr_CheckDir(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) & _
                                     gstr_JobName & DEF_UNDERSCORE & strPressFileName & DEF_UNDERSCORE & SH.Name
                                                                                                                        
        '..get information from each instance of each part in sheet
        For Each Ni In SH.Parts
                                                            
            lngQuantityOnSheet = 0
                                                            
            For Each niPart In SH.Parts

                If niPart.Name = Ni.Name Then lngQuantityOnSheet = lngQuantityOnSheet + 1

            Next niPart
                                            
            Set psNI = Ni.Paths
            
            For Each pNI In psNI
                
                With pNI
                    
                    If .Attribute(DEF_ATT_ALPHADOOR) = "1" Then
                    
                        strPartImage = .Attribute(DEF_ATT_PART_IMAGE)
                        lngItem = CLng(.Attribute(DEF_ATT_ITEM))
                        lngDetailID = CLng(.Attribute(DEF_ATT_DETAIL_ID))
                        lngOrderID = .Attribute(DEF_ATT_ORDER_ID)
                        lngCustomerID = .Attribute(DEF_ATT_CUSTOMER_ID)
                        strType = .Attribute(DEF_ATT_TYPE_NAME)
                            
                        .Attribute(DEF_ATT_PRESS_NAME) = CStr(Press.PressName)
                        .Attribute(DEF_ATT_PRESS_SHEET_NAME) = CStr(SH.Name)
                        .Attribute(DEF_ATT_PRESS_ITEM_NUMBER) = CStr(lngItem)
                        .Attribute(DEF_ATT_PRESS_QTY_ON_SHEET) = CStr(lngQuantityOnSheet)
                        .Attribute(DEF_ATT_PRESS_SHEET_NUMBER) = CStr(lngSheetNumber)
                        .Attribute(DEF_ATT_PRESS_SHEET_NUMBER) = CStr(PressColour.ColourName)
                        
                            
                        Exit For
                        
                    End If
                    
                End With
    
            Next pNI
            
            With colPressReportData("k" & lngDetailID)
              If Not gbln_ItemExistsInCollection(.colSheets, SH.Name) Then
                Set PressReportSheet = New CPressReportSheet
                PressReportSheet.SheetName = SH.Name
                .colSheets.add PressReportSheet, SH.Name
              Else
                Set PressReportSheet = .colSheets(SH.Name)
              End If
            End With
                
            Set PressReportDataItem = New CPressReportDataItem
            PressReportDataItem.TypeName = strType
            PressReportDataItem.QtyOnSheet = lngQuantityOnSheet
            PressReportSheet.colComponents.add PressReportDataItem
            
            sSql = "INSERT INTO AD_REPORT_DATA(DetailID , OrderID, CustomerID, PressName, FoilColour, " & _
              "PressSheetName, PressItemNumber, PressQuantityOnSheet, PressPathToEMF, PressSheetNumber, PathToPressNestARD) " & _
              "VALUES (" & lngDetailID & "," & lngOrderID & "," & glng_CustomerID & ",'" & gs_FixSQL(Press.PressName) & "','" & _
              gs_FixSQL(PressColour.ColourName) & "','" & SH.Name & "'," & lngItem & ",'" & CStr(lngQuantityOnSheet) & "','" & _
              gs_FixSQL(strNestImage) & DEF_EXTENSION_EMF & "'," & lngSheetNumber & ",'" & ActiveDrawing.FullName & "');"
            
            gdb_CDM.Execute sSql
                    
        Next Ni
        
    Next SH
    
    ' 12 Aug 13 - Moved from above to support Alphacam Reports Engine
    ' (Ensures press nesting attributes are saved)
    m_CreateAlphaCAMDrawingsOfSheetsPress strPressFileName
    
    ActiveDrawing.AutoScale ' .ZoomAll
    DoEvents
    
Controlled_Exit:
        
    Set nInfo = Nothing
    Set SH = Nothing
    Set pNI = Nothing

Exit Sub

m_InsertReportData_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "m_InsertReportData"
    MsgBox Err.Description, vbExclamation, Frame.ReadTextFile(strCTX, 120, 1)
    
    Resume Controlled_Exit
    
End Sub

Private Sub m_InsertReportDataRouter(Material As CMaterial)

    Dim rstData                 As ADODB.Recordset
    Dim Ni                      As NestPartInstance
    Dim niPart                  As NestPartInstance
    Dim nInfo                   As NestInformation
    Dim SH                      As NestSheet
    Dim psNI                    As Paths
    Dim psGeos                  As Paths
    Dim pNI                     As Path
    Dim strNestImage            As String
    Dim strPartImage            As String
    Dim strType                 As String
    Dim strStyle                As String
    Dim lngItem                 As Long
    Dim strPartX                As String
    Dim strPartY                As String
    Dim strANCFile              As String
    Dim strANCPath              As String
    Dim strQuantity             As String
    Dim lngQuantityOnSheet      As String
    Dim dblX2                   As Double
    Dim dblY2                   As Double
    Dim dblX1                   As Double
    Dim dblY1                   As Double
    Dim dblScrap                As Double
    Dim lngDetailID             As Long                                             '..07.09.02 - rg
    Dim strCustomerName         As String
    Dim strOrderNumber          As String
    Dim strItemNumber           As String
    Dim lngOrderID              As Long
    Dim sSql                    As String
    Dim sProductionComment      As String
    Dim lngSheetNumber          As Long
    Dim rstReportData           As ADODB.Recordset
    Dim lngPK                   As Long
    Dim strBarcodeID            As String
    Dim strDoorImage            As String
    Dim lngNestDoorCount        As Long
    Dim strFoilColour           As String
    Dim dblNestPartLeft         As Double
    Dim dblNestPartRight        As Double
    Dim dblNestPartTop          As Double
    Dim dblNestPartBottom       As Double
    

On Error GoTo m_InsertReportData_Error

    '..ensure connection
    If Not gbln_ConnectToDB Then GoTo Controlled_Exit

    m_CreateAlphaCAMDrawingsOfSheets Material

    '..assign NInfo to current drawing nest information
    Set nInfo = ActiveDrawing.GetNestInformation

    '..update the process
    With Frame
      .ShowProgressBox .ReadTextFile(strCTX, 120, 1), .ReadTextFile(strCTX, 300, 22) & Space(3)
    End With

    DoEvents

    '..get information from each nested sheet
    For Each SH In nInfo.Sheets

        lngSheetNumber = lngSheetNumber + 1

        '..update the process
        Frame.SetProgressText Frame.ReadTextFile(strCTX, 300, 22) & Space(3) & SH.Name & Space(3)
        DoEvents

        '..set the path to the wmf file
        strNestImage = gstr_CheckDir(gstr_EnsureBackslash(clsOptions.PathToRoot) & DEF_PATH_IMAGE) & _
                                     gstr_JobName & DEF_UNDERSCORE & Material.MaterialName & DEF_UNDERSCORE & SH.Name

        '..get the scrap
        dblScrap = gdbl_Scrap(SH)

        strBarcodeID = gstr_GenerateNestedSheetBarcode
        SH.Path.Attribute(DEF_ATT_SHEET_UNIQUE_ID) = strBarcodeID
        
        '..get information from each instance of each part in sheet
        For Each Ni In SH.Parts

            lngQuantityOnSheet = 0

            For Each niPart In SH.Parts

                If niPart.Name = Ni.Name Then lngQuantityOnSheet = lngQuantityOnSheet + 1

            Next niPart

            Set psNI = Ni.Paths
            Set psGeos = ActiveDrawing.CreatePathCollection
            
            For Each pNI In psNI

                With pNI

                    '..look for a toolpath
                    If .IsToolPath Then
                        If .GetMillData.ProcessType2 <> acamProcessMACHINE_POLYLINE Then

                            strANCFile = .Attribute(DEF_ATT_ANC_NAME)
                            strANCPath = .Attribute(DEF_ATT_ANC_FULLNAME)
                            strPartImage = .Attribute(DEF_ATT_PART_IMAGE)
                            lngItem = CLng(.Attribute(DEF_ATT_ITEM))
                            strDoorImage = .Attribute(DEF_ATT_NEST_DOOR_IMAGE)
                            lngNestDoorCount = .Attribute(DEF_ATT_NEST_DOOR_COUNT)
    
                            '..see if created by AlphaDOOR (must do this for inserted paths)
                            If .Attribute(DEF_ATT_ALPHADOOR) = 1 Then
    
                                strStyle = .Attribute(DEF_ATT_USER_STYLE_NAME)             ' .Attribute(DEF_ATT_STYLE_NUMBER)      '..07.10.02 - rg
                                strType = .Attribute(DEF_ATT_TYPE_NAME)
                                strPartX = IIf(Len(.Attribute(DEF_ATT_PART_WIDTH)) = 0, DEF_NOT_APPLICABLE, .Attribute(DEF_ATT_PART_WIDTH))
                                strPartY = IIf(Len(.Attribute(DEF_ATT_PART_LENGTH)) = 0, DEF_NOT_APPLICABLE, .Attribute(DEF_ATT_PART_LENGTH))
                                strQuantity = .Attribute(DEF_ATT_PART_QUANTITY)
                                lngDetailID = CLng(.Attribute(DEF_ATT_DETAIL_ID))
                                strCustomerName = .Attribute(DEF_ATT_CUST_NAME)
                                strOrderNumber = .Attribute(DEF_ATT_ORDER_NUM)
                                strItemNumber = .Attribute(DEF_ATT_ITEM_NUM)
                                lngOrderID = .Attribute(DEF_ATT_ORDER_ID)
                                sProductionComment = .Attribute(DEF_ATT_DOOR_PRODUCTION_COMMENT)
                                strFoilColour = Material.colDoors("k" & lngDetailID).FoilColour
    
                                Exit For
    
                            Else
    
                                '..assume filler part
                                strStyle = DEF_SCRAP_FILLER
                                strType = Ni.Name
                                .Attribute(DEF_ATT_TYPE_NAME) = Ni.Name
                                strPartX = DEF_NOT_APPLICABLE
                                strPartY = DEF_NOT_APPLICABLE
                                strQuantity = CStr(lngQuantityOnSheet)
    
                                strCustomerName = ""
                                strOrderNumber = ""
                                strItemNumber = ""
                                sProductionComment = ""
    
                                Exit For
    
                            End If
    
                        End If

                    End If
                    
                End With

            Next pNI

            If clsOptions.AdditionalReportData Then
            
                For Each pNI In psNI
                    With pNI
                        If Not .IsToolPath Then
                          If pNI.Attribute(DEF_ATT_GEOMETRY_NUMBER) = "1" Then
                            psGeos.add pNI
                            Exit For
                          End If
                        End If
                    End With
                Next
                
                If psGeos.Count > 0 Then
                  psGeos.GetExtentL dblNestPartLeft, dblNestPartBottom, dblNestPartRight, dblNestPartTop
                  
                  dblNestPartLeft = CDbl(Format(dblNestPartLeft, "0.00"))
                  dblNestPartRight = CDbl(Format(dblNestPartRight, "0.00"))
                  dblNestPartTop = CDbl(Format(dblNestPartTop, "0.00"))
                  dblNestPartBottom = CDbl(Format(dblNestPartBottom, "0.00"))
                  
                End If

            End If
            
            If colReportData.Count > 0 Then
              ' See if this order detail has press data attached
              If gbln_ItemExistsInCollection(colReportData, "k" & lngDetailID) Then
                ' Get the ID of the row of report data
                lngPK = colReportData("k" & lngDetailID).GetNextID
                
                colPressReportData("k" & lngDetailID).UpdatePressInfo lngPK, SH.Name
              
              End If
            End If
    
            ' Test to see if Press Report data exists for this item
            If lngPK = 0 Then
              
                ' This item has not been optimised for the press
                ' Insert the row of data
                sSql = "INSERT INTO AD_REPORT_DATA(DetailID , OrderID, CustomerID, SheetCount," & _
                  "SheetPartCount, SheetLength, SheetMaterial, SheetName, SheetScrap, SheetThickness," & _
                  "SheetWidth, PartItemNumber, PartQuantity, PartQuantityOnSheet, PartType, PartStyle," & _
                  "PartWidth, PartLength, PathToANC, PathToARD, PathToEMF, PathToNestEMF, PartARDName," & _
                  "PartANCName, SheetNumber, ProductionComment, PressSheetIdentifier, PressDoorImage," & _
                  "PressDoorCounter, FoilColour, NestPartPositionLeft, NestPartPositionRight, NestPartPositionTop," & _
                  "NestPartPositionBottom) VALUES (" & lngDetailID & "," & lngOrderID & "," & _
                  glng_CustomerID & ",'" & CStr(nInfo.Sheets.Count) & "','" & CStr(SH.Parts.Count) & "','" & _
                  CStr(clsNest.SheetLength) & "','" & gs_FixSQL(Material.MaterialName) & "','" & SH.Name & "','" & CStr(dblScrap) & "','" & _
                  CStr(clsNest.SheetThickness) & "','" & CStr(clsNest.SheetWidth) & "'," & _
                  lngItem & ",'" & strQuantity & "','" & CStr(lngQuantityOnSheet) & "','" & gs_FixSQL(strType) & "','" & _
                  gs_FixSQL(strStyle) & "','" & strPartX & "','" & strPartY & "','" & gs_FixSQL(strANCPath) & "','" & gs_FixSQL(Ni.FileName) & "','" & _
                  gs_FixSQL(strPartImage) & DEF_EXTENSION_EMF & "','" & gs_FixSQL(strNestImage) & DEF_EXTENSION_EMF & "','" & _
                  gs_FixSQL(Ni.Name) & DEF_EXTENSION_ARD & "','" & gs_FixSQL(strANCFile) & "'," & lngSheetNumber & ",'" & _
                  gs_FixSQL(sProductionComment) & "','" & strBarcodeID & "','" & gs_FixSQL(strDoorImage) & DEF_EXTENSION_EMF & "'," & _
                  lngNestDoorCount & ",'" & gs_FixSQL(strFoilColour) & "'," & gs_NoComma(CStr(dblNestPartLeft)) & "," & _
                  gs_NoComma(CStr(dblNestPartRight)) & "," & gs_NoComma(CStr(dblNestPartTop)) & "," & gs_NoComma(CStr(dblNestPartBottom)) & ");"
              
            Else
                ' This item has been optimised for the press
                ' Edit the row of data with the PK obtained
                
                ' Assign the SQL string
                sSql = "UPDATE AD_REPORT_DATA SET SheetCount='" & CStr(nInfo.Sheets.Count) & "'," & _
                        "SheetPartCount='" & CStr(SH.Parts.Count) & "'," & _
                        "SheetLength='" & CStr(clsNest.SheetLength) & "'," & _
                        "SheetMaterial='" & gs_FixSQL(Material.MaterialName) & "'," & _
                        "SheetName='" & SH.Name & "',SheetScrap='" & CStr(dblScrap) & "'," & _
                        "SheetThickness='" & CStr(clsNest.SheetThickness) & "'," & _
                        "SheetWidth='" & CStr(clsNest.SheetWidth) & "'," & _
                        "PartItemNumber=" & lngItem & "," & _
                        "PartQuantity='" & strQuantity & "'," & _
                        "PartQuantityOnSheet='" & CStr(lngQuantityOnSheet) & "'," & _
                        "PartType='" & gs_FixSQL(strType) & "'," & _
                        "PartStyle='" & gs_FixSQL(strStyle) & "'," & _
                        "PartWidth='" & strPartX & "', PartLength='" & strPartY & "'," & _
                        "PathToANC='" & gs_FixSQL(strANCPath) & "'," & _
                        "PathToARD='" & gs_FixSQL(Ni.FileName) & "'," & _
                        "PathToEMF='" & gs_FixSQL(strPartImage) & DEF_EXTENSION_EMF & "'," & _
                        "PathToNestEMF='" & gs_FixSQL(strNestImage) & DEF_EXTENSION_EMF & "'," & _
                        "PartARDName ='" & gs_FixSQL(Ni.Name) & DEF_EXTENSION_ARD & "'," & _
                        "PartANCName='" & gs_FixSQL(strANCFile) & "'," & _
                        "SheetNumber=" & lngSheetNumber & ", ProductionComment='" & gs_FixSQL(sProductionComment) & "'," & _
                        "PressSheetIdentifier='" & strBarcodeID & "', PressDoorImage='" & gs_FixSQL(strDoorImage) & DEF_EXTENSION_EMF & "'," & _
                        "PressDoorCounter=" & lngNestDoorCount & ", " & _
                        "NestPartPositionLeft=" & gs_NoComma(CStr(dblNestPartLeft)) & ",NestPartPositionRight=" & gs_NoComma(CStr(dblNestPartRight)) & " ,NestPartPositionTop=" & gs_NoComma(CStr(dblNestPartTop)) & ", NestPartPositionBottom=" & gs_NoComma(CStr(dblNestPartBottom)) & " " & _
                        "WHERE PK=" & lngPK
              
            End If
            
            ' Execute the SQL
            gdb_CDM.Execute sSql

        Next Ni

    Next SH

    ActiveDrawing.AutoScale ' .ZoomAll
    DoEvents





Controlled_Exit:

    If Not (rstData Is Nothing) Then
        With rstData
            If (.State = adStateOpen) Then .Close
        End With
    End If

    Set nInfo = Nothing
    Set SH = Nothing
    Set pNI = Nothing

Exit Sub

m_InsertReportData_Error:

    If (Err.Number <> 0) Then WriteError Err, True, "m_InsertReportData"
    MsgBox Err.Description, vbExclamation, Frame.ReadTextFile(strCTX, 120, 1)

    If Not (rstData Is Nothing) Then
        With rstData
            If (.EditMode = adEditAdd Or adEditInProgress) Then .CancelUpdate
        End With
    End If

    Resume Controlled_Exit

End Sub




Private Sub m_FillOrderInfo(DoorComponent As CDoor, ByVal rOrder As ADODB.Recordset, sJob As String, lJob As Long)
    
    Dim rstCustomer             As ADODB.Recordset
    
    With DoorComponent
        
        .OrderID = lJob
        .JobName = sJob
        .ProcessedDate = Now
        
        '..order info
        .CustomerID = gvar_CheckNull(rOrder.Fields!CustomerID)
        .HotJob = CBool(gvar_CheckNull(rOrder.Fields!HotJob))
        .DueDate = gvar_CheckNull(rOrder.Fields!DueDate)
        .OrderDate = gvar_CheckNull(rOrder.Fields!OrderDate)
        .PO = gvar_CheckNull(rOrder.Fields!PO)
        
        Set rstCustomer = grst_GetCustomerInfo(.CustomerID)
        
        If (rstCustomer Is Nothing) Then
        
            '..customer info
            .Address_1 = vbNullString
            .Address_2 = vbNullString
            .City = vbNullString
            .Contact = vbNullString
            .CustomerName = vbNullString
            .Email = vbNullString
            .Fax = vbNullString
            .Telephone = vbNullString
            .Zip = vbNullString
        
        Else
        
            '..customer info
            .Address_1 = gvar_CheckNull(rstCustomer.Fields!Address_1)
            .Address_2 = gvar_CheckNull(rstCustomer.Fields!Address_2)
            .City = gvar_CheckNull(rstCustomer.Fields!City)
            .Contact = gvar_CheckNull(rstCustomer.Fields!Contact)
            .CustomerName = gvar_CheckNull(rstCustomer.Fields!Name)
            .Email = gvar_CheckNull(rstCustomer.Fields!Email)
            .Fax = gvar_CheckNull(rstCustomer.Fields!Fax)
            .Telephone = gvar_CheckNull(rstCustomer.Fields!Telephone)
            .Zip = gvar_CheckNull(rstCustomer.Fields!Zip)
            
            If (rstCustomer.State = adStateOpen) Then rstCustomer.Close
            Set rstCustomer = Nothing
            
        End If
                
    End With

End Sub

Private Sub m_SplitNestANC(Material As CMaterial, nstThisNest As NestInformation, _
                          strThisFilePath As String, strThisFileName As String, colANC As Collection, OutputSecondLocation As Boolean)

    Dim nstSheets               As NestSheets
    Dim NstSheet                As NestSheet
    Dim colSheetNames           As Collection
    Dim nFile                   As Integer
    Dim nFiles                  As Integer
    Dim nSheetCount             As Integer
    Dim strLine                 As String
    Dim strNewFileName          As String
    Dim strNewFilePath          As String
    Dim START_FLAG              As String
    Dim END_FLAG                As String
    Dim intFreeFile             As Integer
    Dim blnSTART                As Boolean
    Dim strControlFile          As String
    Dim strAutoLoadFile         As String
    Dim FSO                     As New Scripting.FileSystemObject
    Dim strPallet1Material      As String
    Dim strPallet2Material      As String
    Dim strPallet               As String

On Error GoTo PROC_ERR
    
    ' -- Get all the sheet names
    Set nstSheets = nstThisNest.Sheets
    
    Set colSheetNames = New Collection
    
    For Each NstSheet In nstSheets
        
        '..add to collection
        'colSheetNames.Add Replace$(Trim$(nstSheet.Name), " ", "_")
    
        ' ** LJO
        colSheetNames.add DetermineSheetID(NstSheet.Name)
    
    Next
    
    ' -- Set counter
    nSheetCount = 1
    
    strControlFile = gstr_EnsureBackslash(gstr_ParseDir(strThisFilePath)) & "Worklist.seq"
    strAutoLoadFile = gstr_EnsureBackslash(gstr_ParseDir(strThisFilePath)) & "Pallet.txt"
    
    ' If the control file exists, delete it.
    If FSO.FileExists(strControlFile) Then
      FSO.DeleteFile strControlFile
    End If
    
    ' If the auto load file exists, delete it.
    If FSO.FileExists(strAutoLoadFile) Then
      FSO.DeleteFile strAutoLoadFile
    End If
    
    With clsOptions
    
        ' -- Retrieve settings from reg
        START_FLAG = .SplitProgramsStartFlag
        END_FLAG = .SplitProgramsEndFlag
         
'        strExtension = .SplitProgramsExtension
        
        ' -- Open the FileName and scan for each sheet
        nFile = FreeFile
        
        '..lool for start
        blnSTART = mbln_HasSTART(strThisFilePath)
        
        ' -- Open the nc file
        Open strThisFilePath For Input As #nFile
        
        ' -- Scan each line
        Do While Not EOF(nFile)
            
            ' -- Get a line
            Line Input #nFile, strLine
            
            strLine = Trim$(strLine)
            
            ' -- Make sure we have something to test
            If Len(strLine) > 0 Then
                                               
                ' -- Check for start of file (at beginning of line)
                If Left$(strLine, Len(START_FLAG)) = START_FLAG Then
                'InStr(1, strLine, START_FLAG, vbTextCompare) > 0 Then
                                       
                    ' -- Start a new file
                    nFiles = FreeFile
                    
                    ' -- Start a new nc file
                    strNewFileName = strThisFileName & DEF_UNDERSCORE & colSheetNames(nSheetCount) & "." & clsOptions.NCFileExtension
                    
                    If OutputSecondLocation Then
                    
                        If clsOptions.OutputSecondCreateSubFolder Then
                          strNewFilePath = gstr_EnsureBackslash(clsOptions.PathToRootSecond) & gstr_JobName & "\" & strNewFileName
                        Else
                          strNewFilePath = gstr_EnsureBackslash(clsOptions.PathToRootSecond) & strNewFileName
                        End If
                      
                    Else
                      
                        If clsOptions.OutputResultsSubFolder Then
                          strNewFilePath = gstr_EnsureBackslash(clsOptions.PathToRoot) & gstr_JobName & "\" & strNewFileName
                        Else
                          strNewFilePath = gstr_EnsureBackslash(clsOptions.PathToRoot) & strNewFileName
                        End If
                      
                    End If
                    
                    Open strNewFilePath For Output As #nFiles
                    
                    ' -- Write to the file
                    
                    '..reinsert START?
                    If blnSTART Then
                        Print #nFiles, "START"
                    End If
                    
                    '..print the start of file marker
                    If Len(Trim$(.SplitProgramsStartOfFileMarker)) > 0 Then
                        Print #nFiles, .SplitProgramsStartOfFileMarker
                    End If
                    
                    Print #nFiles, strLine
                    
                    ' -- Scan each line
                    Do While Not EOF(nFile)
                        
                        ' -- Stay in this loop until eof marker is found
                        Line Input #nFile, strLine

                        If InStr(1, strLine, END_FLAG, vbTextCompare) > 0 Then
                            
                            ' -- We are done with this file
                            Print #nFiles, strLine
                            
                            '..print the end of file marker
                            If Len(Trim$(.SplitProgramsEndOfFileMarker)) > 0 Then
                                Print #nFiles, .SplitProgramsEndOfFileMarker
                            End If

                            Exit Do
                                                                            
                        Else
                        
                            ' -- Write to the file
                            Print #nFiles, strLine
                        
                        End If
                            
                    Loop
                
                    ' -- Close the new nc file
                    Close #nFiles
                    
                    ' -- See if we need to convert to PGM
                    If clsOptions.SplitProgramsOutputPGM Then
                                                
                        ' -- Convert all the new files
                        ' -- Convert file
                        Call m_ConvertToPGM(strNewFilePath)

                        ' -- Check resulting inf files to see if there were any errors raised
                        ' -- Check INF and delete if required
                        Call m_CheckForErrors(strNewFilePath)
                                                      
                    End If
                                                           
                    '..update the nest anc path in the database
                    If Not clsOptions.DisableReports Then
                    
                        Call m_UpdateDBNestPaths(DEF_RST_RPT_NESTANC, DEF_RST_RPT_PATHTONESTANC, strNewFilePath, strNewFileName, True, nSheetCount)
                    
                    End If
                    
                    If clsOptions.SplitControlFile Then
                      intFreeFile = FreeFile
                      Open strControlFile For Append As #intFreeFile
                        Print #intFreeFile, gstr_ParseName(strNewFilePath)
                      Close #intFreeFile
                    End If
                    
                    If clsOptions.SplitAutoLoad Then
                      strPallet1Material = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_AUTOLOAD, DEF_REG_KEY_AUTOLOAD_PALLET_1, "")
                      strPallet2Material = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_AUTOLOAD, DEF_REG_KEY_AUTOLOAD_PALLET_2, "")
                      strPallet = ""
                      
                      If UCase(strPallet1Material) = UCase(Material.MaterialName) Then
                        strPallet = "1"
                      End If
                      
                      If UCase(strPallet2Material) = UCase(Material.MaterialName) Then
                        strPallet = "2"
                      End If
                      
                      intFreeFile = FreeFile
                      
                      Open strAutoLoadFile For Append As #intFreeFile
                        Print #intFreeFile, gstr_PadString(gstr_ParseName(strNewFilePath), 79) & strPallet
                      Close #intFreeFile
                    
                    End If
                    
                    colANC.add strNewFilePath
                    
                    nSheetCount = nSheetCount + 1

                End If

            End If
            
        Loop
        
    End With
    
    Close #nFile ' -- Close the nc file
      
PROC_EXIT:

     ' -- Clean up
     Set NstSheet = Nothing
     Set nstSheets = Nothing
     Set colSheetNames = Nothing

Exit Sub

PROC_ERR:

    MsgBox Err.Description, vbExclamation, Err.Source
    Resume PROC_EXIT

End Sub

Private Function mbln_HasSTART(strThisFile As String) As Boolean

    Dim nFile                   As Integer
    Dim strLine                 As String

On Error GoTo PROC_ERR
    
    mbln_HasSTART = False

    ' -- Get a free file handle
    nFile = FreeFile
    
    ' -- overwrite the nc file
    Open strThisFile For Input As #nFile
        
    ' -- Scan each line
    Do While Not EOF(nFile)
    
        ' -- Get a line
        Line Input #nFile, strLine
       
        ' -- Make sure we have something to test
        If Len(strLine) > 0 Then
            
            If (UCase$(Mid$(strLine, 1, 5)) = "START") Then
                
                mbln_HasSTART = True
                Exit Do
            
            End If
                            
        End If
                        
    Loop
    
    Close #nFile

Exit Function

PROC_ERR:

    MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
    mbln_HasSTART = False
    Close

End Function

Private Sub m_ConvertToPGM(strFile As String)
    
    Dim FSO                     As New Scripting.FileSystemObject
    Dim strPath                 As String
    Dim strTempFile             As String

    ' -- Conversion executable
    Const WINXISO               As String = "winxiso.exe"

On Error GoTo PROC_ERR

Debug.Print strFile

    ' -- Assign Paths
    strPath = gstr_CheckDir(App.Frame.PathOfThisAddin) & WINXISO
    
    ' -- Check for exe
    If Not FSO.FileExists(strPath) Then
                                    
       MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 104) & Space(1) & strPath, vbExclamation, App.Name
       Exit Sub
       
    End If
    
    strTempFile = FSO.GetFile(strFile).ShortPath
    
    ShellAndWait strPath & Space(1) & strTempFile & " -s -i"

PROC_EXIT:
    
    Set FSO = Nothing

Exit Sub

PROC_ERR:
                            
    MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
    Resume PROC_EXIT

End Sub

Private Sub m_CheckForErrors(strFilename As String)

On Error GoTo PROC_ERR
    
    Dim FSO                     As New Scripting.FileSystemObject
    Dim strINF                  As String
    Dim strError                As String
    Dim intFreeFile             As Integer

    ' -- Assign filename
    strINF = FSO.GetFile(strFilename).ShortPath
    
    ' -- Change ext to match inf output
    Mid$(strINF, Len(strINF) - 2, 3) = "INF"
    
    ' -- Make sure it exists
    If Len(Dir$(strINF, vbNormal)) > 0 Then
       
       ' -- Get a free handle to a file
       intFreeFile = FreeFile
       
       ' -- Open the INF file
       Open strINF For Input As #intFreeFile
            
            ' -- Loop to the last line
            Do While Not EOF(intFreeFile): Line Input #intFreeFile, strError: Loop
            
       Close #intFreeFile
     
     End If
         
    ' -- Check for any errors
    If Right$(strError, 1) = "0" Then
       
        ' -- Are we killing files ?
        If clsOptions.SplitProgramsPurgeINF Then
            ' -- No errors so kill this file and original [xxl]
            FSO.DeleteFile strINF, True
        End If
       
    Else
       ' -- [ERRORS]=1
       MsgBox strINF & vbCrLf & App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 105) & Space(1) & _
              Mid$(strError, InStr(1, strError, "=") + 1, 999) & Space(1) & _
              App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 106), vbExclamation, DEF_PROJECT_NAME
                                    
    End If
               
PROC_EXIT:
    
    Set FSO = Nothing
               
Exit Sub
PROC_ERR:

    MsgBox Err.Description, vbExclamation, Err.Source
    Resume PROC_EXIT
               
End Sub

Private Function mpths_PathsInGroup(pPaths As Paths, iGroup As Integer) As Paths        '..04.10.02 - rg
    
    Dim pthPath                 As Path
    Dim pthsPaths               As Paths
    
On Error GoTo mpths_PathsInGroup_Error
    
    Set pthsPaths = ActiveDrawing.CreatePathCollection
    
    For Each pthPath In pPaths
        If pthPath.Group = iGroup Then pthsPaths.add pthPath
    Next pthPath

    Set mpths_PathsInGroup = pthsPaths
    
Controlled_Exit:
    
    Set pthPath = Nothing
    Set pthsPaths = Nothing
    
Exit Function

mpths_PathsInGroup_Error:
    
    MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
    If (Err.Number <> 0) Then WriteError Err, True, "mpths_PathsInGroup"
    Set mpths_PathsInGroup = Nothing
    Resume Controlled_Exit
    
End Function

Private Function mpths_PathsNotInGroup(pPaths As Paths, iOffsetFrom As Integer) As Paths     '..04.10.02 - rg
    
    Dim pthPath                 As Path
    Dim pthsPaths               As Paths
    
On Error GoTo mpths_PathsNotInGroup_Error
    
    Set pthsPaths = ActiveDrawing.CreatePathCollection

    For Each pthPath In pPaths
        If pthPath.Attribute(DEF_ATT_GEOMETRY_NUMBER) = iOffsetFrom Then
            pthsPaths.add pthPath
            Exit For
        End If
    Next pthPath

    Set mpths_PathsNotInGroup = pthsPaths
    
Controlled_Exit:

    Set pthPath = Nothing
    Set pthsPaths = Nothing
    
Exit Function

mpths_PathsNotInGroup_Error:
    
    MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
    If (Err.Number <> 0) Then WriteError Err, True, "mpths_PathsNotInGroup"
    Set mpths_PathsNotInGroup = Nothing
    Resume Controlled_Exit
        
End Function

Private Function mbln_MachineWithOffset(cPath As CPathData, pthPath As Path, pthPick As Path, _
                                        pthBoundary As Path, blnHasBoundary As Boolean) As Boolean
    
    Dim pthsOffset              As Paths
    Dim pthsBoundary            As Paths
    Dim pthPartial              As Path
    
On Error GoTo mbln_MachineWithOffset_Error
    
    mbln_MachineWithOffset = True
    
    With cPath
    
        If pthPath.Closed Then
            
            '..check the direction of the geo to be offset
            If pthPath.CW Then
                                                                                        
                '..check the side to offset
                If .PathOffsetSide = acamINSIDE Then
                                                                                                                                        
                    '..inside
                    Set pthsOffset = pthPath.Offset(.PathOffsetValue, acamRIGHT)
                                                                                                                    
                    '..set the path to be machined
                    Set pthPick = pthsOffset(1)
                
                    
                    If .CutType <> DEF_CUT_TYPE_PARTIAL Then
                    
                        '..set the start point
                        If Not gbln_SetStartPoint(pthPick, .LeadEntryPointIsCorner, False) Then
                        
                            '..couldn't do it for some reason
                            MsgBox Frame.ReadTextFile(strCTX, 500, 23) & Space(3), _
                                   vbInformation, Frame.ReadTextFile(strCTX, 120, 1)
                        
                        End If
                
                    Else
                    
                        
                        Set pthPartial = DrawParametricPartialPath(pthPick, .PartialStartElemIndex, .PartialStartElemDist, .PartialEndElemIndex, .PartialEndElemDist, .ToolSidePartialReverse)
                        pthPick.Selected = True
                        ActiveDrawing.DeleteSelected
                        
                        Set pthPick = pthPartial
                    
                    End If
                
                Else
                    
                    '..outside
                    Set pthsOffset = pthPath.Offset(.PathOffsetValue, acamLEFT)
                    
                    '..set the path to be machined
                    Set pthPick = pthsOffset(1)
                
                    If .CutType <> DEF_CUT_TYPE_PARTIAL Then
                    
                        '..set the start point
                        If Not gbln_SetStartPoint(pthPick, .LeadEntryPointIsCorner, False) Then
                        
                            '..couldn't do it for some reason
                            MsgBox Frame.ReadTextFile(strCTX, 500, 23) & Space(3), _
                                   vbInformation, Frame.ReadTextFile(strCTX, 120, 1)
                        
                        End If
                    
                    Else
                    
                        Set pthPartial = DrawParametricPartialPath(pthPick, .PartialStartElemIndex, .PartialStartElemDist, .PartialEndElemIndex, .PartialEndElemDist, .ToolSidePartialReverse)
                        pthPick.Selected = True
                        ActiveDrawing.DeleteSelected
                        
                        Set pthPick = pthPartial
                    
                    End If
                
                End If
                
                '..now look for a boundary
                If (.PocketBoundary <> 0) Then
                    
                    '..offset the new geo to the inside
                    Set pthsBoundary = pthPick.Offset(.PocketBoundary - (.ToolDiameter / 2), acamRIGHT)
                    
                    '..create the new geometry
                    Set pthBoundary = pthsBoundary(1)
                    
                    '..make it a different color
                    pthBoundary.Color = acamDARK_RED
                    
                    '..let it know that there's a boundary
                    blnHasBoundary = True
                                                            
                End If
                                            
            Else
            
                '..check for the side to offset
                If .PathOffsetSide = acamINSIDE Then
                    
                    '..inside
                    Set pthsOffset = pthPath.Offset(.PathOffsetValue, acamLEFT)
                    
                    '..create the path to be machined
                    Set pthPick = pthsOffset(1)
                    
                    If .CutType <> DEF_CUT_TYPE_PARTIAL Then
                    
                        '..set the start point
                        If Not gbln_SetStartPoint(pthPick, .LeadEntryPointIsCorner, False) Then
                        
                            '..couldn't do it for some reason
                            MsgBox Frame.ReadTextFile(strCTX, 500, 23) & Space(3), _
                                   vbInformation, Frame.ReadTextFile(strCTX, 120, 1)
                        
                        End If
                    
                    Else
                                                
                        Set pthPartial = DrawParametricPartialPath(pthPick, .PartialStartElemIndex, .PartialStartElemDist, .PartialEndElemIndex, .PartialEndElemDist, .ToolSidePartialReverse)
                        pthPick.Selected = True
                        ActiveDrawing.DeleteSelected
                        
                        Set pthPick = pthPartial
                    
                    End If
                    
                    
                Else
                    
                    '..outside
                    Set pthsOffset = pthPath.Offset(.PathOffsetValue, acamRIGHT)
                    
                    '..create the path to be machined
                    Set pthPick = pthsOffset(1)
                    
                    If .CutType <> DEF_CUT_TYPE_PARTIAL Then
                    
                        '..set the start point
                        If Not gbln_SetStartPoint(pthPick, .LeadEntryPointIsCorner, False) Then
                        
                            '..couldn't do it for some reason
                            MsgBox Frame.ReadTextFile(strCTX, 500, 23) & Space(3), _
                                   vbInformation, Frame.ReadTextFile(strCTX, 120, 1)
                        
                        End If
    
                    Else
                    
                        Set pthPartial = DrawParametricPartialPath(pthPick, .PartialStartElemIndex, .PartialStartElemDist, .PartialEndElemIndex, .PartialEndElemDist, .ToolSidePartialReverse)
                        pthPick.Selected = True
                        ActiveDrawing.DeleteSelected
                        
                        Set pthPick = pthPartial
                    
                    End If
    
                End If
                
                '..now look for a boundary
                If (.PocketBoundary <> 0) Then
                    
                    '..offset the new geo to the inside
                    Set pthsBoundary = pthPick.Offset(.PocketBoundary - (.ToolDiameter / 2), acamLEFT)
                    
                    '..create the new geometry
                    Set pthBoundary = pthsBoundary(1)
                    
                    '..make it a different color
                    pthBoundary.Color = acamDARK_RED
                    
                    '..let it know that there's a boundary
                    blnHasBoundary = True
                                                            
                End If
                                    
            End If
        
        Else
        
            '..open geo always in clockwise motion by default
            If .PathOffsetSide = acamINSIDE Then
                Set pthsOffset = pthPath.Offset(.PathOffsetValue, acamRIGHT)
            Else
                Set pthsOffset = pthPath.Offset(.PathOffsetValue, acamLEFT)
            End If
            
            Set pthPick = pthsOffset(1)
        
            If .CutType = DEF_CUT_TYPE_PARTIAL Then
            
                Set pthPartial = DrawParametricPartialPath(pthPick, .PartialStartElemIndex, .PartialStartElemDist, .PartialEndElemIndex, .PartialEndElemDist, .ToolSidePartialReverse)
                pthPick.Selected = True
                ActiveDrawing.DeleteSelected
                
                Set pthPick = pthPartial
            
            End If
        
        End If
    
    End With
    
    '..make it a different color than the original
    pthPick.Color = acamBROWN
    
    DoEvents

Controlled_Exit:

    Set pthsOffset = Nothing
    Set pthsBoundary = Nothing

Exit Function

mbln_MachineWithOffset_Error:
    
    MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
    mbln_MachineWithOffset = False
    Resume Controlled_Exit

End Function

Private Function mbln_MachineWithNoOffset(cPath As CPathData, pthPick As Path, _
                                          pthBoundary As Path, blnHasBoundary As Boolean) As Boolean
    
    Dim pthsBoundary            As Paths
    
On Error GoTo mbln_MachineWithNoOffset_Error
    
    mbln_MachineWithNoOffset = True

    With cPath

        '..set the start point
        If Not gbln_SetStartPoint(pthPick, .LeadEntryPointIsCorner, False) Then
        
            '..couldn't do it for some reason
            MsgBox Frame.ReadTextFile(strCTX, 500, 23) & Space(3), _
                   vbInformation, Frame.ReadTextFile(strCTX, 120, 1)
        
        End If
        
        If (.PocketBoundary <> 0) Then
            
            '..check direction so we know what side to offset
            If pthPick.CW Then
            
                '..offset the new geo to the inside
                Set pthsBoundary = pthPick.Offset(.PocketBoundary - (.ToolDiameter / 2), acamRIGHT)
            
            Else
                
                '..offset the new geo to the inside
                Set pthsBoundary = pthPick.Offset(.PocketBoundary - (.ToolDiameter / 2), acamLEFT)
            
            End If
            
            '..create the new geometry
            Set pthBoundary = pthsBoundary(1)
            
            '..make it a different color
            pthBoundary.Color = acamDARK_RED
            
            '..let it know that there's a boundary
            blnHasBoundary = True
        
        End If
    
    End With

Controlled_Exit:

    Set pthsBoundary = Nothing

Exit Function

mbln_MachineWithNoOffset_Error:
    
    MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
    mbln_MachineWithNoOffset = False
    Resume Controlled_Exit

End Function

Private Sub m_SetDetailAttributes(Door As CDoor, P As Path)
    
On Error Resume Next
    
    With Door

        P.Attribute("LicomUSrlg_alphadoor_CustomerName") = .CustomerName
        P.Attribute("LicomUSrlg_alphadoor_JobID") = .JobName
        P.Attribute("LicomUSrlg_alphadoor_PO") = .PO
        P.Attribute("LicomUSrlg_alphadoor_Address_1") = .Address_1
        P.Attribute("LicomUSrlg_alphadoor_Address_2") = .Address_2
        P.Attribute("LicomUSrlg_alphadoor_City") = .City
        P.Attribute("LicomUSrlg_alphadoor_Zip") = .Zip
        P.Attribute("LicomUSrlg_alphadoor_Telephone") = .Telephone
        P.Attribute("LicomUSrlg_alphadoor_Fax") = .Fax
        P.Attribute("LicomUSrlg_alphadoor_Contact") = .Contact
        P.Attribute("LicomUSrlg_alphadoor_Email") = .Email
        P.Attribute("LicomUSrlg_alphadoor_DueDate") = .DueDate
        P.Attribute("LicomUSrlg_alphadoor_OrderDate") = .OrderDate
        P.Attribute(DEF_ATT_TYPE_NAME) = .TypeName
        P.Attribute("LicomUSrlg_alphadoor_HotJob") = CStr(.HotJob)
        P.Attribute("LicomUSrlg_alphadoor_StyleNumber") = .StyleNumber
        P.Attribute("LicomUSrlg_alphadoor_Quantity") = .Quantity
        P.Attribute(DEF_ATT_PART_WIDTH) = .Width
        P.Attribute(DEF_ATT_PART_LENGTH) = .Length
        P.Attribute("LicomUSrlg_alphadoor_CornerRadius") = .CornerRadius
        P.Attribute("LicomUSrlg_alphadoor_UserStyleName") = .UserStyleName
        P.Attribute("LicomUSrlg_alphadoor_UserVariableString") = .UserVariableString
        P.Attribute("LicomUSrlg_alphadoor_UserArg_0") = .UserArg_0
        P.Attribute("LicomUSrlg_alphadoor_UserArg_1") = .UserArg_1
        P.Attribute("LicomUSrlg_alphadoor_UserArg_2") = .UserArg_2
        P.Attribute("LicomUSrlg_alphadoor_UserArg_3") = .UserArg_3
        P.Attribute("LicomUSrlg_alphadoor_UserArg_4") = .UserArg_4
        P.Attribute("LicomUSrlg_alphadoor_UserArg_5") = .UserArg_5
        P.Attribute("LicomUSrlg_alphadoor_UserArg_6") = .UserArg_6
        P.Attribute(DEF_ATT_DOOR_PRODUCTION_COMMENT) = Door.ProductionComment
        P.Attribute(DEF_ATT_DOOR_CUSTOM_1) = Door.CustomField1
        P.Attribute(DEF_ATT_DOOR_CUSTOM_2) = Door.CustomField2
        P.Attribute(DEF_ATT_FOIL_COLOUR) = Door.FoilColour
        
        P.Attribute(DEF_ATT_GROUP_ID) = P.Group
        
    End With
    
    If (Err.Number <> 0) Then Err.Clear
    
End Sub

Private Sub m_SetPressAttributes(Door As CDoor, P As Path)
    
On Error Resume Next
    
    With Door

        P.Attribute(DEF_ATT_ALPHADOOR) = "1"
        P.Attribute(DEF_ATT_DETAIL_ID) = .DetailID
        P.Attribute(DEF_ATT_ORDER_ID) = .OrderID
        P.Attribute(DEF_ATT_CUSTOMER_ID) = .CustomerID
        P.Attribute(DEF_ATT_TYPE_NAME) = Door.TypeName
        P.Attribute(DEF_ATT_FOIL_COLOUR) = Door.FoilColour
        
    End With
    
    If (Err.Number <> 0) Then Err.Clear
    
End Sub


Private Function mbln_Style_900(Door As CDoor, lngPartNumber As Long, colANC As Collection, bPreview As Boolean, bPress As Boolean, bStopProcessing As Boolean)
    
    Dim FSO                     As New Scripting.FileSystemObject
    Dim pVol                    As Path
    Dim lngGeoNumber            As Long
    
On Error GoTo mbln_Style_900_Error

    '..start out ok
    mbln_Style_900 = True
    
    '..open the inserted drawing
    With Door
    
        '..make sure the insert file exists
        If Not FSO.FileExists(.TypeName) Then mbln_Style_900 = False: GoTo Controlled_Exit
    
        If Not mbln_Intialize(lngGeoNumber) Then mbln_Style_900 = False: GoTo Controlled_Exit
        
        '..open the file to insert and validate it
        App.OpenDrawing .TypeName
        
    End With
    
    If bPress Then
        
        ' Remove toolpaths
        ActiveDrawing.SetToolPathsSelected True
        ActiveDrawing.DeleteSelected
    
        Call m_SetPressAttributes(Door, ActiveDrawing.Geometries(1))
        
        '..making for the Press - just save the door without any machining
        '..update the database and save the file
        mbln_Style_900 = mbln_UpdateAndSavePress(Door, lngPartNumber)
    
    Else
    
        '..look for toolpaths
        If ActiveDrawing.ToolPaths.Count = 0 Then
            
            With App.Frame
            
                MsgBox .ReadTextFile(clsOptions.CTXFile, 500, 95) & vbCrLf & _
                       .ReadTextFile(clsOptions.CTXFile, 500, 96) & vbCrLf & vbCrLf & _
                       UCase$(Door.TypeName) & Space(3), vbInformation, DEF_PROJECT_NAME
                       
                '..set to success so no error during the run
                mbln_Style_900 = False: GoTo Controlled_Exit
            
            End With
            
        Else
            
            '..attach some attributes
            Call m_SetAttributes(ActiveDrawing.ToolPaths(1), Door) ' uDTD.TypeName, uDTD.Width, uDTD.Length)       '..07.09.02 - rg
            
        End If
            
        '..loop thru toolpaths looking for invalid depths
        For Each pVol In ActiveDrawing.ToolPaths
        
            With pVol
                If Not gbln_CheckDepthTolerance(.GetMillData.MaterialTop, .GetMillData.FinalDepth) Then
                    '..set to success so no error during the run
                    mbln_Style_900 = False: GoTo Controlled_Exit
                End If
            End With
        
        Next pVol
            
        '..clean out any work volume if exists
        For Each pVol In ActiveDrawing.Geometries
             
            With pVol
                .Selected = .IsWorkVolume
            End With
            
        Next pVol
                    
        With ActiveDrawing
        
            '..delete the selected (should be the work volume)
            .DeleteSelected
    
            '..let's see the damn thing
            '.ZoomAll
            '.Redraw
            
        End With
                       
        '..reasign to the type name to the file name and replace extension decimal
        With clsTypeData
            .TypeName = gstr_ParseName(Door.TypeName)
            .TypeName = Replace$(.TypeName, ".", "_")
        End With
    
        Dim dblMinX As Double
        Dim dblMinY As Double
        Dim dblMaxX As Double
        Dim dblMaxY As Double
        
        ActiveDrawing.Geometries.GetExtentL dblMinX, dblMinY, dblMaxX, dblMaxY
        
        Door.Width = dblMaxX - dblMinX
        Door.Length = dblMaxY - dblMinY
        
        If Not mbln_DrawHandleHoles(Door, bPreview) Then
            
            mbln_Style_900 = False
            bStopProcessing = True
            GoTo Controlled_Exit
        
        End If
        
        ' TFS #56993 - Run a custom macro
        If clsOptions.CustomMacro <> "" Then
          If Not mbln_RunCustomMacro(clsOptions.CustomMacro, Door) Then
            mbln_Style_900 = False
            GoTo Controlled_Exit
          End If
        End If
        
        If Not mbln_UpdateAndSave(Door, lngPartNumber, colANC, bPreview) Then
            
            mbln_Style_900 = False
            GoTo Controlled_Exit
            
        End If
    End If


Controlled_Exit:

    '..rinse
    Set FSO = Nothing
    
Exit Function
    
mbln_Style_900_Error:
          
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_Style_900"
    mbln_Style_900 = False
    Resume Controlled_Exit

End Function

Private Function mbln_Style_910(Door As CDoor, lngPartNumber As Long, colANC As Collection, bPreview As Boolean)
    
    Dim FSO                     As New Scripting.FileSystemObject
    Dim pVol                    As Path
    Dim lngGeoNumber            As Long
    
On Error GoTo mbln_Style_910_Error

    '..start out ok
    mbln_Style_910 = True
    
    '..open the inserted drawing
    With Door
    
        '..make sure the insert file exists
        If Not FSO.FileExists(.TypeName) Then mbln_Style_910 = False: GoTo Controlled_Exit
    
        If Not mbln_Intialize(lngGeoNumber) Then mbln_Style_910 = False: GoTo Controlled_Exit

        Frame.CloseProgressBox
    
        '..attempt to run the macro
        App.RunParametricMacro .TypeName

        DoEvents
        
    End With
    
    '..look for toolpaths
    If ActiveDrawing.ToolPaths.Count = 0 Then
        
        With Frame
        
            MsgBox .ReadTextFile(clsOptions.CTXFile, 500, 95) & vbCrLf & _
                   .ReadTextFile(clsOptions.CTXFile, 500, 96) & vbCrLf & vbCrLf & _
                   UCase$(Door.TypeName) & Space(3), vbInformation, DEF_PROJECT_NAME
                   
            '..set to success so no error during the run
            mbln_Style_910 = False: GoTo Controlled_Exit
        
        End With
    
    Else
        
        '..attach some attributes
        Call m_SetAttributes(ActiveDrawing.ToolPaths(1), Door)
                    
    End If
        
    '..loop thru toolpaths looking for invalid depths
    For Each pVol In ActiveDrawing.ToolPaths
    
        With pVol
            If Not gbln_CheckDepthTolerance(.GetMillData.MaterialTop, .GetMillData.FinalDepth) Then
                '..set to success so no error during the run
                mbln_Style_910 = False: GoTo Controlled_Exit
            End If
        End With
    
    Next pVol
        
    '..clean out any work volume if exists
    For Each pVol In ActiveDrawing.Geometries
         
        With pVol
            .Selected = .IsWorkVolume
        End With
        
    Next pVol
                
    With ActiveDrawing
    
        '..delete the selected (should be the work volume)
        .DeleteSelected

        '..let's see the damn thing
        '.ZoomAll
        '.Redraw
        
    End With
                   
    '..reasign to the type name to the file name and replace extension decimal
    With clsTypeData
        .TypeName = gstr_ParseName(Door.TypeName)
        .TypeName = Replace$(.TypeName, ".", "_")
    End With

    '..do the machining
    If Not mbln_UpdateAndSave(Door, lngPartNumber, colANC, bPreview) Then
        
        mbln_Style_910 = False
        GoTo Controlled_Exit
        
    End If

Controlled_Exit:

    '..rinse
    Set FSO = Nothing
    
Exit Function
    
mbln_Style_910_Error:
          
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_Style_910"
    mbln_Style_910 = False
    Resume Controlled_Exit

End Function

Private Function mbln_Style_920(Door As CDoor, lngPartNumber As Long, colANC As Collection, bPreview As Boolean)
    
    Dim FSO                     As New Scripting.FileSystemObject
    Dim pVol                    As Path
    Dim lngGeoNumber            As Long
    
On Error GoTo mbln_Style_920_Error

    '..start out ok
    mbln_Style_920 = True
    
    '..open the inserted drawing
    With Door
    
        '..make sure the insert file exists
        If Not gbln_ProjectExists(.TypeName) Then
            MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 98) & Space(1) & UCase$(.TypeName) & _
                   App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 99) & Space(3), vbInformation, DEF_PROJECT_NAME
            mbln_Style_920 = False
            GoTo Controlled_Exit
        End If
            
        If Not mbln_Intialize(lngGeoNumber) Then mbln_Style_920 = False: GoTo Controlled_Exit

        App.Frame.CloseProgressBox
    
        '..attempt to run the macro
        Call App.Run(.TypeName & "." & DEF_ADOOR_MOD & "." & DEF_ADOOR_PROC)

        DoEvents
        
    End With
    
    '..look for toolpaths
    If ActiveDrawing.ToolPaths.Count = 0 Then
        
        With App.Frame
        
            MsgBox .ReadTextFile(clsOptions.CTXFile, 500, 95) & vbCrLf & _
                   .ReadTextFile(clsOptions.CTXFile, 500, 96) & vbCrLf & vbCrLf & _
                   UCase$(Door.TypeName) & Space(3), vbInformation, DEF_PROJECT_NAME
                   
            '..set to success so no error during the run
            mbln_Style_920 = False: GoTo Controlled_Exit
        
        End With
            
    Else
        
        '..attach some attributes
        Call m_SetAttributes(ActiveDrawing.ToolPaths(1), Door)
            
    End If
        
    '..loop thru toolpaths looking for invalid depths
    For Each pVol In ActiveDrawing.ToolPaths
    
        With pVol
            If Not gbln_CheckDepthTolerance(.GetMillData.MaterialTop, .GetMillData.FinalDepth) Then
                '..set to success so no error during the run
                mbln_Style_920 = False: GoTo Controlled_Exit
            End If
        End With
    
    Next pVol
        
    '..clean out any work volume if exists
    For Each pVol In ActiveDrawing.Geometries
         
        With pVol
            .Selected = .IsWorkVolume
        End With
        
    Next pVol
                
    With ActiveDrawing
    
        '..delete the selected (should be the work volume)
        .DeleteSelected

        '..let's see the damn thing
        '.ZoomAll
        '.Redraw
        
    End With
                   
    '..reasign to the type name to the file name and replace extension decimal
    With clsTypeData
        .TypeName = gstr_ParseName(Door.TypeName)
        .TypeName = Replace$(.TypeName, ".", "_")
    End With
    
    '..do the machining
    If Not mbln_UpdateAndSave(Door, lngPartNumber, colANC, bPreview) Then
        
        mbln_Style_920 = False
        GoTo Controlled_Exit
        
    End If

Controlled_Exit:

    '..rinse
    Set FSO = Nothing
    
Exit Function
    
mbln_Style_920_Error:
          
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_Style_920"
    mbln_Style_920 = False
    Resume Controlled_Exit

End Function
'EditMark
Private Function mbln_Style_Make_930(Door As CDoor, bPress As Boolean, Optional lngPartNumber As Long, _
                                     Optional colANC As Collection, Optional bPreview As Boolean)
      
      
      
'    EditMark
    Dim pVol                    As Path
    Dim pthOut                  As Path
    Dim pthUser                 As Path
    Dim i                       As Integer
    Dim iL                      As Integer
    Dim iU                      As Integer
    Dim strDims()               As String
    Dim strDimsDesc()           As String
    Dim lngGeoNumber            As Long
    Dim RequiredData            As CUserStyle
    Dim vbaProject              As CVBAProject
    Dim bVBAProjectOpen         As Boolean
    Dim dll                     As Object
    
On Error GoTo mbln_Style_Make_930_Error

    '..start out ok
    mbln_Style_Make_930 = True
    
    Set RequiredData = New CUserStyle
    
    '..open the inserted drawing
    With Door
                
        '..make sure the vba project is loaded
        If Not gbln_ProjectExists(.UserStyleName) Then
            MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 98) & Space(1) & UCase$(.UserStyleName) & Space(1) & _
                   App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 99) & Space(3), vbInformation, DEF_PROJECT_NAME
            mbln_Style_Make_930 = False
            GoTo Controlled_Exit
        End If
            
        If Not mbln_Intialize(lngGeoNumber) Then mbln_Style_Make_930 = False: GoTo Controlled_Exit
    
        '..assign dimensions
        RequiredData.Width = .Width
        RequiredData.Length = .Length
        RequiredData.CornerRadius = .CornerRadius
        
        RequiredData.Success = True
    
        '..split the dimension string
        strDims = Split(.UserVariableString, ";")
        
        iL = LBound(strDims)
        iU = UBound(strDims)
    
        For i = iL To iU
            RequiredData.UserVariables.add PDbl(Trim$(strDims(i)))
        Next i

        DoEvents
                
        '..split the dimension description string
        strDimsDesc = Split(.UserVariableDescriptionString, ";")
        
        iL = LBound(strDimsDesc)
        iU = UBound(strDimsDesc)
        
        For i = iL To iU
            RequiredData.UserVariableDescriptions.add Trim$(strDimsDesc(i))
        Next i
        
        DoEvents
                
        '..skipping outside geo?                                                    '..07.03.02 - rg
        If Not .IgnoreOuterGeometry Then
                    
            '..create the door perimeter
            
          Set pthOut = ActiveDrawing.CreateRectangle(0, 0, RequiredData.Width, RequiredData.Length)
          Set dll = CreateObject("StdAlpha.ShareClass")
          dll.jc pthOut, RequiredData.UserVariables(46), RequiredData.UserVariables(47), RequiredData.UserVariables(48), RequiredData.UserVariables(49)
          dll.diamond pthOut, .CornerRadius
          pthOut.Group = 0

                                                                 '..07.22.02 - rg
            
            'With ActiveDrawing
            '    .ZoomAll
            '    .Redraw
            'End With
            
            DoEvents
        
            '..assign the geometry number (first outside pass should always be 1)
            lngGeoNumber = (lngGeoNumber + 1)
            pthOut.Attribute(DEF_ATT_GEOMETRY_NUMBER) = CStr(lngGeoNumber)
            
            If bPress Then
              Call m_SetPressAttributes(Door, pthOut)
            Else
              Call m_SetDetailAttributes(Door, pthOut)
            End If
        End If
    
        ' Get the correct VBAProject object
        Set vbaProject = colVBAUserStyles(.UserStyleName)
        
        ' Test to see if the project is already open within VBA
        ' Loading an already open project will reset any breakpoint the user may have set
        bVBAProjectOpen = gbln_IsVBAProjectOpen(vbaProject.ProjectName)
        
        If Not bVBAProjectOpen Then
          ' Load the VBA User Style
          App.LoadAddIn vbaProject.FileName
        End If
                
        '..launch the desired macro and create geometry
        Call App.Run(.UserStyleName & "." & DEF_ADOOR_MOD & "." & DEF_ADOOR_PROC, RequiredData, _
                     .UserArg_0, .UserArg_1, .UserArg_2, .UserArg_3, .UserArg_4, .UserArg_5, .UserArg_6)
    
        If Not bVBAProjectOpen Then
          ' Finished with the VBA Project, so disable it (unload it)
          App.EnableAddIn vbaProject.FileName, False
        End If
    
    End With
      
    If RequiredData.PathsToReturn Is Nothing Then mbln_Style_Make_930 = False: GoTo Controlled_Exit
    If Not RequiredData.Success Then mbln_Style_Make_930 = False: GoTo Controlled_Exit

    For Each pthUser In RequiredData.PathsToReturn
        
        ' Add press attributes to first user style generated geometry
        ' to overcome issue if VBA macro deletes outer path :(
        If lngGeoNumber = 1 And bPress Then
          m_SetPressAttributes Door, pthUser
        End If
        
        lngGeoNumber = (lngGeoNumber + 1)
        pthUser.Attribute(DEF_ATT_GEOMETRY_NUMBER) = CStr(lngGeoNumber)
        
        Call m_SetDetailAttributes(Door, pthUser)
        
    Next pthUser
                    
    '..clean out any work volume if exists
    For Each pVol In ActiveDrawing.Geometries
         
        With pVol
            .Selected = .IsWorkVolume
        End With
        
    Next pVol
                
    With ActiveDrawing
    
        '..delete the selected (should be the work volume)
        .DeleteSelected
        '.ZoomAll
        '.Redraw
        DoEvents
        
    End With
                   
    If Not bPress Then
                           
        '..do the machining
        If Not mbln_MakeMachining(Door, lngGeoNumber, lngPartNumber, colANC, bPreview) Then
            
            mbln_Style_Make_930 = False
            GoTo Controlled_Exit
            
        End If

    Else
            
        '..making for the Press - just save the door without any machining
        '..update the database and save the file
        mbln_Style_Make_930 = mbln_UpdateAndSavePress(Door, lngPartNumber)
    
    End If
    If RequiredData.UserVariables(50) <> 0 Then
    dll.bz RequiredData.UserVariables(50)
    End If
    Set dll = Nothing

Controlled_Exit:

    '..rinse
    Erase strDims
    
    Set pthOut = Nothing
    Set pthUser = Nothing
    Set RequiredData = Nothing

Exit Function
    
mbln_Style_Make_930_Error:
          
    If Err.Number = -2147467259 Then
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 100) & Space(3), vbExclamation + vbMsgBoxHelpButton, _
                DEF_PROJECT_NAME, App.Frame.PathOfThisAddin & DEF_BACKSLASH & DEF_HELP_NAME_CHM, DEF_HLP_ID_Defining_a_User_Defined_Door_Style
    Else
        If (Err.Number <> 0) Then WriteError Err, True, "mbln_Style_Make_930"
    End If
    
    mbln_Style_Make_930 = False
    Resume Controlled_Exit

End Function

Public Sub Temp()

  Dim lngOpCount  As Long
  Dim Op          As Operation
  Dim Ops         As Operations
  
  Set Ops = ActiveDrawing.Operations
  
  For lngOpCount = 1 To Ops.Count
    Set Op = Ops(lngOpCount)
    Set Op = Nothing
  Next

End Sub

Public Sub TestScrap()
  Dim Ni As NestInformation
  Dim Ns As NestSheet
'
  Set Ni = ActiveDrawing.GetNestInformation
  For Each Ns In Ni.Sheets
    Debug.Print gdbl_Scrap(Ns)
  Next
End Sub


