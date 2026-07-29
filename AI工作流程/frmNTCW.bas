




















Option Explicit

Private Sub m_SetInsertDrawingControls()
  chkInsertFileCenterX.Enabled = Not optRefParametric.Value
  chkInsertFileCenterY.Enabled = Not optRefParametric.Value
  txtParametricGroupNumber.Enabled = optRefParametric.Value
  lblInsertPointX.Enabled = Not optRefParametric.Value
  lblInsertPointY.Enabled = Not optRefParametric.Value
  txtDistanceFromRefX.Enabled = Not optRefParametric.Value
  txtDistanceFromRefY.Enabled = Not optRefParametric.Value
  fraInsertPositionPoint.Enabled = Not optRefParametric.Value
End Sub


Private Function mbln_AdvanceDoorCreation(iWiz As AdoorWizardStage) As Boolean

On Error GoTo Error
    
    '..ok so far
    mbln_AdvanceDoorCreation = True
    
    With clsPathData
    
        '..set the machining method
        .CreationMethod = Switch(optManualDefine, DEF_CREATION_METHOD_MANUAL, _
                                 optMachiningStyle, DEF_CREATION_METHOD_MACHINING_STYLE)
                
        '..now lauch the next step depending on machine method
        Select Case .CreationMethod
    
            '..manually define?
            Case DEF_CREATION_METHOD_MANUAL: iWiz = adoorWIZARD_MACHINE_METHOD
                                        
            '..machining style?
            Case Else:
                
                ' Test to ensure some machining styles have been defined
                If MillMachiningStyles.Count > 0 Then
                    iWiz = adoorWIZARD_OFFSET_AMOUNT
                Else
                    MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 73, 4) & Space(2), vbExclamation, DEF_PROJECT_NAME
                    mbln_AdvanceDoorCreation = False
                End If
                
        End Select
    
    End With
    
Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceDoorCreation"
    mbln_AdvanceDoorCreation = False
    Call g_UnLoadAllForms

End Function

Private Function mbln_AdvancePocketLeads() As Boolean
    Dim Ctl             As Control
'
On Error GoTo Error
    
    '..start out ok
    mbln_AdvancePocketLeads = True
    
    '..make sure all data is entered
    If Not mbln_CheckAllText(Ctl) Then
    
        mbln_AdvancePocketLeads = False
        GoTo Controlled_Exit
        
    End If
    
    If Not mbln_SetLeadDataPocket Then
        
        '..let the user know there's a problem
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 120) & Space(3), vbExclamation, DEF_PROJECT_NAME
               
        mbln_AdvancePocketLeads = False
        
        GoTo Controlled_Exit
    End If
    
    If chkUse3DApproach And Val(txtApproachAngle) = 0 Then
        MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 121) & Space(3), vbExclamation, DEF_PROJECT_NAME
        
        mbln_AdvancePocketLeads = False
        
        GoTo Controlled_Exit
    End If
    
    If chkUse3DApproach And Val(txtZigZagLength) = 0 Then
        MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 122) & Space(3), vbExclamation, DEF_PROJECT_NAME
        
        mbln_AdvancePocketLeads = False
        
        GoTo Controlled_Exit
    End If
        
    DoEvents

    '..hide me and select geometry
    Me.Hide
    
Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvancePocketLeads"
    mbln_AdvancePocketLeads = False
    Call g_UnLoadAllForms


End Function

Private Sub chkDepthPercentage_Change()

    Dim bVisible                As Boolean

On Error Resume Next

    bVisible = chkDepthPercentage.Value
    
    txtFinalDepth.Visible = Not bVisible
    txtFinalDepthPercent.Visible = bVisible
    
    txtThicknessOfFirstCut.Visible = Not bVisible
    txtThicknessOfFirstCutPercent.Visible = bVisible
    
    txtThicknessOfLastCut.Visible = Not bVisible
    txtThicknessOfLastCutPercent.Visible = bVisible
    
    lblPercent1.Visible = bVisible
    lblPercent2.Visible = bVisible
    lblPercent3.Visible = bVisible
    
End Sub

Private Sub chkInsertFileCenterY_Click()

On Error Resume Next

    '..enable and disable controls as necessary
    If chkInsertFileCenterY.Value Then
        
        optRefBottomLeft.Enabled = False
        optRefBottomRight.Enabled = False
        optRefTopLeft.Enabled = False
        optRefTopRight.Enabled = False
        
        optRefBottomLeft.Visible = False
        optRefBottomRight.Visible = False
        optRefTopLeft.Visible = False
        optRefTopRight.Visible = False
        
        optRefRight.Visible = True
        optRefLeft.Visible = True
        optRefTop.Visible = True
        optRefBottom.Visible = True
        
        If chkInsertFileCenterX.Value = True Then
            
            fraInsertPositionPoint.Enabled = False
            lblInsertPointX.Enabled = False
            lblInsertPointY.Enabled = False
            txtDistanceFromRefX.Enabled = False
            txtDistanceFromRefY.Enabled = False
            
            optRefRight.Enabled = False
            optRefLeft.Enabled = False
            optRefTop.Enabled = False
            optRefBottom.Enabled = False
                            
        Else
        
            fraInsertPositionPoint.Enabled = True
            lblInsertPointX.Enabled = True
            lblInsertPointY.Enabled = False
            txtDistanceFromRefX.Enabled = True
            txtDistanceFromRefY.Enabled = False
            
            optRefRight.Enabled = True
            optRefLeft.Enabled = True
            optRefTop.Enabled = False
            optRefBottom.Enabled = False
            
            optRefLeft.Value = True
                            
        End If
        
    Else
        
        If chkInsertFileCenterX.Value = True Then
            
            optRefBottomLeft.Enabled = False
            optRefBottomRight.Enabled = False
            optRefTopLeft.Enabled = False
            optRefTopRight.Enabled = False
            
            optRefBottomLeft.Visible = False
            optRefBottomRight.Visible = False
            optRefTopLeft.Visible = False
            optRefTopRight.Visible = False
            
            optRefRight.Visible = True
            optRefLeft.Visible = True
            optRefTop.Visible = True
            optRefBottom.Visible = True
            
            optRefRight.Enabled = False
            optRefLeft.Enabled = False
            optRefTop.Enabled = True
            optRefBottom.Enabled = True
            
            optRefTop.Value = True
            
            fraInsertPositionPoint.Enabled = True
            lblInsertPointX.Enabled = False
            lblInsertPointY.Enabled = True
            txtDistanceFromRefX.Enabled = False
            txtDistanceFromRefY.Enabled = True
            
        Else
                    
            optRefBottomLeft.Visible = True
            optRefBottomRight.Visible = True
            optRefTopLeft.Visible = True
            optRefTopRight.Visible = True
                        
            optRefBottomLeft.Enabled = True
            optRefBottomRight.Enabled = True
            optRefTopLeft.Enabled = True
            optRefTopRight.Enabled = True
            
            optRefTopLeft.Value = True
            
            optRefRight.Enabled = False
            optRefLeft.Enabled = False
            optRefTop.Enabled = False
            optRefBottom.Enabled = False
            
            optRefRight.Visible = False
            optRefLeft.Visible = False
            optRefTop.Visible = False
            optRefBottom.Visible = False
            
            fraInsertPositionPoint.Enabled = True
            lblInsertPointX.Enabled = True
            lblInsertPointY.Enabled = True
            txtDistanceFromRefX.Enabled = True
            txtDistanceFromRefY.Enabled = True
            
        End If

    End If
    
    If (Err.Number <> 0) Then WriteError Err, False, "chkInsertFileCenterY_Click"
    
End Sub

Private Sub chkSlowDownForCorners_Click()
    
  g_EnableFrame fraSlowDownForCorners, chkSlowDownForCorners

End Sub

Private Sub chkUse3DApproach_Click()
  fra3DApproachParameters.Enabled = chkUse3DApproach
  lblApproachAngle.Enabled = chkUse3DApproach
  txtApproachAngle.Enabled = chkUse3DApproach
  lblZigZagLength.Enabled = chkUse3DApproach
  txtZigZagLength.Enabled = chkUse3DApproach
End Sub

Private Sub cmdBack_Click()

On Error GoTo Error
        
        DoEvents
        
        '..find wizard stage
        Call m_WizardBack
            
        DoEvents
        
Controlled_Exit:

Exit Sub

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "frmWizard_cmdBack_Click"
    Call g_UnLoadAllForms
    
End Sub

Private Sub cmdBackplot_Click()
        
On Error GoTo ErrTrap
    
        Me.Hide
    
        With App.ActiveDrawing
            
                ' 02 mar 11 TFS#44393
                '   + MODIFIED to run simulation
                '
                .ThreeDViews = True
                .ZoomAll
                DoEvents
                                                
                '.Options.ShowTools = True
                '.Redraw
                'DoEvents
                '.Options.ShowTools = False
                '
                App.Frame.RunCommand acamCmdVIEW_SIMULATION
                        
                DoEvents
                .ThreeDViews = False
                .ZoomAll
        
        End With
    
Controlled_Exit:
                
        DoEvents
        Me.Show



Exit Sub

ErrTrap:

        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 45) & Space(3), vbExclamation, DEF_PROJECT_NAME
        If (Err.Number <> 0) Then WriteError Err, True, "cmdBackplot_Click"
        Resume Controlled_Exit
   
End Sub

Private Sub cmdHelp_Click()
    
    LoadHelp

End Sub

Private Sub cmdNext_Click()
    
On Error GoTo Error
        
        DoEvents
        
        '..find wizard stage
        Call m_WizardNext
        
Controlled_Exit:

Exit Sub

Error:
    
    With App.Frame
    
        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "frmWizard_cmdNext_Click"
    Call g_UnLoadAllForms

End Sub

Private Sub cmdCancel_Click()
        
On Error Resume Next
    
    DoEvents
    
    '..ask to leave the wizard
    If MsgBox(App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 47) & Space(3), _
              vbQuestion + vbYesNo, DEF_PROJECT_NAME) = vbNo Then Exit Sub
                  
    With clsWizard                                                                      '..07.02.02 - rg
                                   
        '..any paths applied
        If clsPathData.OperationNumber > .AddedOrginalNumber Then
        
            '..save the current door type name to the registry for the vb side to grab
            SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_TYPE, DEF_KEY_ACTIVE_TYPE, clsTypeData.TypeName
        
            '..now ask to remove applied paths
            If MsgBox(App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 48) & Space(3), _
                      vbQuestion + vbYesNo, DEF_PROJECT_NAME) = vbYes Then
                                    
                '..didn't make it so let the dll know
                SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_CANCELWIZARD, -1
                                    
                .RemoveType = True
                    
            Else
    
                '..cancel but still get defaults
                SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_CANCELWIZARD, 0
            
                .RemoveType = False
                
            End If
                
        Else
        
            '..didn't make it so let the dll know
            SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_CANCELWIZARD, -1
            
            .RemoveType = True
            
        End If
    
        .WizardCanceled = True
        .AddAnotherPath = False

    End With
    
Controlled_Exit:
                
    '..clear drawing
    App.New
    
    Me.Hide
    Unload Me
        
    If (Err.Number <> 0) Then WriteError Err, False, "frmWizard_cmdCancel_Click"
        
End Sub

Private Sub chkViewGhostTools_Click()
    With App.ActiveDrawing
        .Options.ShowGhostTools = chkViewGhostTools.Value
        .Redraw
        DoEvents
    End With
End Sub


Private Sub MultiPage1_Change()

End Sub

Private Sub optCompBoth_Click()

    chkApplyCompOnRapid.Enabled = True
    
End Sub

Private Sub optOffsetSideCenter_Click()
    
    lblOffset.Enabled = False
    txtOffset.Enabled = False
    txtOffset.Text = "0"
    
    If (Err.Number <> 0) Then WriteError Err, False, "optOffsetSideCenter_Click"

End Sub

Private Sub optRefBottomLeft_Click()
  m_SetInsertDrawingControls
End Sub

Private Sub optRefBottomRight_Click()
  m_SetInsertDrawingControls
End Sub


Private Sub optRefParametric_Click()
  m_SetInsertDrawingControls
End Sub

Private Sub optRefTopLeft_Click()
  m_SetInsertDrawingControls
End Sub


Private Sub optRefTopRight_Click()
  m_SetInsertDrawingControls
End Sub


Private Sub SpinButton1_Change()

On Error Resume Next
    
    txtNumberOfCuts.Text = SpinButton1.Value

End Sub

Private Sub txtChordError_Enter()
    With txtChordError
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtComments_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

On Error Resume Next

    Select Case KeyAscii
      Case 39               'apostrophe
        KeyAscii = 96       'backwards apostrophe
      Case Else
        KeyAscii = KeyAscii
    End Select

End Sub

Private Sub txtCutDirection_Enter()
    With txtCutDirection
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCutFeed_Enter()
    With txtCutFeed
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDistanceFromRefX_Enter()
    With txtDistanceFromRefX
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDistanceFromRefY_Enter()
    With txtDistanceFromRefY
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtDownFeed_Enter()
    With txtDownFeed
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEngravingMaxAngle_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtEngravingMaxAngle, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtEngravingMaxAngle_BeforeUpdate"

End Sub

Private Sub txtEngravingMaxAngle_Enter()
    With txtEngravingMaxAngle
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtFinalDepth_Enter()
    With txtFinalDepth
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFinalDepthPercent_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    '..eval value
    gbln_TextCalc txtFinalDepthPercent, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtFinalDepthPercent_BeforeUpdate"

End Sub

Private Sub txtFinalDepthPercent_Enter()
    With txtFinalDepthPercent
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLeadApproachAngle_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtLeadApproachAngle, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtLeadApproachAngle_BeforeUpdate"

End Sub

Private Sub txtLeadApproachAngle_Enter()
    With txtLeadApproachAngle
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLeadArcRadius_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtLeadArcRadius, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtLeadArcRadius_BeforeUpdate"

End Sub

Private Sub txtLeadArcRadius_Enter()
    With txtLeadArcRadius
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLeadLineLengthIn_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    '..eval value
    gbln_TextCalc txtLeadLineLengthIn, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtLeadLineLengthIn_BeforeUpdate"

End Sub

Private Sub txtLeadLineLengthIn_Enter()
    With txtLeadLineLengthIn
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLeadLineLengthOut_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    '..eval value
    gbln_TextCalc txtLeadLineLengthOut, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtLeadLineLengthOut_BeforeUpdate"

End Sub

Private Sub txtLeadLineLengthOut_Enter()
    With txtLeadLineLengthOut
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLeadOverlap_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtLeadOverlap, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtLeadOverlap_BeforeUpdate"

End Sub

Private Sub txtLeadOverlap_Enter()
    With txtLeadOverlap
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMaterialTop_Enter()
    With txtMaterialTop
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNewTypeCreatedBy_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

On Error Resume Next

    Select Case KeyAscii
      Case 39               'apostrophe
        KeyAscii = 96       'backwards apostrophe
      Case Else
        KeyAscii = KeyAscii
    End Select

End Sub

Private Sub txtNewTypeName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

On Error Resume Next

    Select Case KeyAscii
      Case 39               'apostrophe
        KeyAscii = 96       'backwards apostrophe
      Case Else
        KeyAscii = KeyAscii
    End Select

End Sub

Private Sub txtNumberOfCuts_Enter()
    With txtNumberOfCuts
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtOffset_Enter()
    With txtOffset
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtOffsetNumber_Enter()
    With txtOffsetNumber
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPocketBoundary_Enter()
    With txtPocketBoundary
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtRapidDownTo_Enter()
    With txtRapidDownTo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSafeRapidLevel_Enter()
    With txtSafeRapidLevel
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSpindleSpeed_Enter()
    With txtSpindleSpeed
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtStepLength_Enter()
    With txtStepLength
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtStockToBeLeft_Enter()
    With txtStockToBeLeft
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtThicknessOfFirstCut_Enter()
    With txtThicknessOfFirstCut
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtThicknessOfFirstCutPercent_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim dblDepth                As Double
    
On Error Resume Next
    
    If Not chkDepthPercentage.Value Then Exit Sub
    
    '..eval value
    If gbln_TextCalc(txtThicknessOfFirstCutPercent, Cancel) Then
        
        With txtThicknessOfFirstCutPercent
    
            If PDbl(.Text) < 0 Then
                Beep
                Cancel = True
                
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            
                Exit Sub
            End If

            If PDbl(txtNumberOfCuts.Text) = 2 Then
            
                dblDepth = (100 - PDbl(.Text))
                txtThicknessOfLastCutPercent.Text = gs_NoComma(CStr(dblDepth))
    
            End If
        
        End With
        
    End If

    If (Err.Number <> 0) Then WriteError Err, False, "txtThicknessOfFirstCutPercent_BeforeUpdate"

End Sub

Private Sub txtThicknessOfFirstCutPercent_Enter()
    With txtThicknessOfFirstCutPercent
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtThicknessOfLastCut_Enter()
    With txtThicknessOfLastCut
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtThicknessOfLastCutPercent_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim dblDepth                As Double
    
On Error Resume Next
    
    If Not chkDepthPercentage.Value Then Exit Sub
    
    '..eval value
    If gbln_TextCalc(txtThicknessOfLastCutPercent, Cancel) Then
    
        With txtThicknessOfLastCutPercent
        
            If PDbl(.Text) < 0 Then
                Beep
                Cancel = True
                
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            
                Exit Sub
            End If
        
            If PDbl(txtNumberOfCuts.Text) = 2 Then
            
                dblDepth = (100 - PDbl(.Text))
                txtThicknessOfFirstCutPercent.Text = gs_NoComma(CStr(dblDepth))
            
            End If
            
        End With
           
    End If
   
    If (Err.Number <> 0) Then WriteError Err, False, "txtThicknessOfLastCutPercent_BeforeUpdate"

End Sub

Private Sub txtThicknessOfLastCutPercent_Enter()
    With txtThicknessOfLastCutPercent
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtToolDiameter_Enter()
    With txtToolDiameter
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtToolNumber_Enter()
    With txtToolNumber
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWidthOfCut_Enter()
    With txtWidthOfCut
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub UserForm_Activate()

    '..close the little box
    App.Frame.CloseProgressBox

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'    Dim iFile               As Integer
'    Dim sFile               As String
'    Dim Ctl                 As Control
'    Dim Ctls                As Control
'
'On Error Resume Next
'
'    Const DEF_COMMA As String = ","
'
'    iFile = FreeFile
'
'    sFile = App.Frame.PathOfThisAddin & "\WizNames.txt"
'
'    Open sFile For Output As iFile
'
''    Print #iFile, "NAME,TYPE" ',ENABLED,VALUE,TEXTALIGN,TABINDEX,HEIGHT,WIDTH"
'
'    For Each Ctl In Me.Controls
'
'        With Ctl
'
'            If TypeOf Ctl Is MSForms.Frame Then
'
'                Print #iFile, "-- " & Ctl.Name & " : " & Ctl.Caption & " --"
'
'                For Each Ctls In Ctl.Controls
'
'                    If TypeOf Ctls Is TextBox Then
'
'                        Print #iFile, Ctls.Name   '& DEF_COMMA & .Type '& DEF_COMMA & .Enabled & DEF_COMMA & .Value & _
'                          DEF_COMMA & .TextAlign & DEF_COMMA & .TabIndex & .Height & .Width
'
'                    Else
'
'                        Print #iFile, Ctls.Name & " : " & Ctls.Caption  '& DEF_COMMA & .Type '& DEF_COMMA & .Enabled & DEF_COMMA & .Value & _
'                          DEF_COMMA & .TextAlign & DEF_COMMA & .TabIndex & .Height & .Width
'
'                    End If
'
'                Next Ctls
'
'            End If
'
'        End With
'
''        If Not TypeOf Ctl Is TextBox Then Print #iFile, Ctl.Name
'
'    Next Ctl
'
'    Close

End Sub

Private Sub UserForm_Initialize()

    Dim FSO                 As New Scripting.FileSystemObject
    Dim sRegSection         As String
    Dim sDefault            As String
    
On Error GoTo Initialize_Error

    If Not mbln_FillFormCaptions Then GoTo Initialize_Error
    
    With clsWizard

        '..are starting new or added path
        If .AddedPath Then

            '..added path so start at the machining method
            .WizardStage = adoorWIZARD_CREATION_METHOD  ' WIZARD_MACHININGMETHOD

        Else

            '..new type so start at the beginning
            .WizardStage = adoorWIZARD_DOOR_TYPE ' WIZARD_DOORTYPE

        End If

    End With

    '..hide the < Back button
    cmdBack.Enabled = False

'    With lblWizardStage
'        .Left = 0
'        .Width = Me.Width
'    End With
'
'    With lblWizardStageDescription
'        .Left = 0
'        .Width = Me.Width
'    End With
'
'    With lblOperationNumber
'        .Left = 0
'    End With
'
'    With lblToolName
'        .Left = 0
'    End With
    
    With Label2
        .Left = -2
        .Width = .Width + 22
    End With

    '..defaults
    If CBool(GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, "Units", 0)) Then

        '..metric values
        txtDownFeed.Text = "2500"
        txtCutFeed.Text = "7500"
        txtMaterialTop.Text = "0"
        txtFinalDepth.Text = "0"
        txtFinalDepthPercent.Text = "0"

    Else

        '..inch values
        txtDownFeed.Text = "100"
        txtCutFeed.Text = "300"
        txtMaterialTop.Text = "0"
        txtFinalDepth.Text = "0"
        txtFinalDepthPercent.Text = "0"

    End If
    
    '..get last lead values
    txtLeadArcRadius.Text = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, DEF_REG_KEY_LEADARC, "0")
    txtLeadLineLengthIn.Text = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, DEF_REG_KEY_LEADLINE, "2")         '..07.18.02 - rg
    txtLeadLineLengthOut.Text = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, DEF_REG_KEY_LEADLINEOUT, "2")     '..07.18.02 - rg
    
    '..get the proper registry key according to alphacam level
    Select Case App.ProgramLevel
        Case acamLevelADVANCED: sRegSection = DEF_REG_SECTION_MACHINING_ADV
        Case acamLevelADVANCED3D3AXIS: sRegSection = DEF_REG_SECTION_MACHINING_3AX
        Case acamLevelADVANCED3D5AXIS: sRegSection = DEF_REG_SECTION_MACHINING_5AX
    End Select
    
    '..get the default safe rapid level from the registry
    sDefault = RegistryGetKeyValue(rrkHKeyCurrentUser, sRegSection, DEF_REG_KEY_SAFE)
                
    If Len(Trim$(sDefault)) = 0 Then
        txtSafeRapidLevel.Text = "0"
    Else
        sDefault = Format$(CStr(sDefault), "#0.0000")
        txtSafeRapidLevel.Text = gstr_StripTrailingZeros(gs_NoComma(sDefault))
    End If

    '..get the default rapid down to level from the registry
    sDefault = RegistryGetKeyValue(rrkHKeyCurrentUser, sRegSection, DEF_REG_KEY_RAPID)
                
    If Len(Trim$(sDefault)) = 0 Then
        txtRapidDownTo.Text = "0"
    Else
        sDefault = Format$(CStr(sDefault), "#0.0000")
        txtRapidDownTo.Text = gstr_StripTrailingZeros(gs_NoComma(sDefault))
    End If

    '..setup the spin button
    With SpinButton1

        .Min = 1
        .Max = 100
        .Value = 1

    End With
    
    With MultiPage1
        
        .Style = fmTabStyleNone
        .Left = -1
        .Width = Me.Width + 1
                
    End With

Controlled_Exit:

    '..set the font sizes
    Call g_SetAccelerators(Me)

    Set FSO = Nothing

Exit Sub

Initialize_Error:

    With App.Frame
        MsgBox Err.Description, vbExclamation, .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 49)
    End With

    If (Err.Number <> 0) Then WriteError Err, True, "frmWizard_Initialize"
    Call g_UnLoadAllForms

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

On Error Resume Next

    If CloseMode <> 1 Then Call cmdCancel_Click
                
End Sub

Private Sub UserForm_Terminate()
    
On Error Resume Next

    '..save lead vals
    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, DEF_REG_KEY_LEADARC, txtLeadArcRadius.Text
    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, DEF_REG_KEY_LEADLINE, txtLeadLineLengthIn.Text       '..07.18.02 - rg
    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, DEF_REG_KEY_LEADLINEOUT, txtLeadLineLengthOut.Text   '..07.18.02 - rg
    
    '..make sure the little box goes away
    App.Frame.CloseProgressBox
    
    Set frmNTCW = Nothing

End Sub

Private Sub m_WizardNext()
    
    Dim iWizardStage            As AdoorWizardStage
    Dim strTool                 As String
    
On Error GoTo m_WizardNext_Error

    With clsWizard
    
        '..find the wizard stage
        Select Case .WizardStage
    
            Case adoorWIZARD_DOOR_TYPE
                
                If mbln_AdvanceDoorType Then .WizardStage = adoorWIZARD_CREATION_METHOD
                        
            Case adoorWIZARD_CREATION_METHOD
            
                If mbln_AdvanceDoorCreation(iWizardStage) Then .WizardStage = iWizardStage
            
            Case adoorWIZARD_MACHINE_METHOD
            
                If mbln_AdvanceMachiningMethod(iWizardStage) Then .WizardStage = iWizardStage

            Case adoorWIZARD_INSERT_PATH
                
                If mbln_AdvanceInsertPath Then .WizardStage = adoorWIZARD_PATH_COMPLETE
            
            Case adoorWIZARD_OFFSET_AMOUNT
            
                If mbln_AdvanceOffsetAmount Then .WizardStage = adoorWIZARD_TOOL_DIRECTION_SIDE
            
            Case adoorWIZARD_TOOL_DIRECTION_SIDE
                
                ' 23 sep 11 TFS#46506
                '
                If Not (App.GetCurrentTool Is Nothing) Then strTool = App.GetCurrentTool.FileName
                
                '..hide me so we can get a tool
                DoEvents
                Me.Hide
                
                If mbln_AdvanceToolDirectionSide(iWizardStage) Then .WizardStage = iWizardStage
                                                               
                ' 23 sep 11 TFS#46506
                '
                If Not (App.GetCurrentTool Is Nothing) Then
                        .ToolInfoSet = (StrComp(App.GetCurrentTool.FileName, strTool, vbTextCompare) = 0)
                End If
                                                               
                DoEvents
                                                                                           
                '..bring me back
                Me.Show
                                               
            Case adoorWIZARD_POCKETING
                
                If mbln_AdvancePocket Then .WizardStage = adoorWIZARD_MACHINING_1
            
            Case adoorWIZARD_ROUGH_FINISH
            
                If mbln_AdvanceRoughFinish Then .WizardStage = adoorWIZARD_MACHINING_1
                    
            Case adoorWIZARD_MACHINING_1
                    
                If mbln_AdvanceMachiningOptions_1 Then .WizardStage = adoorWIZARD_MACHINING_2
                    
            Case adoorWIZARD_MACHINING_2
                    
                If mbln_AdvanceMachiningOptions_2(iWizardStage) Then .WizardStage = iWizardStage
            
            Case adoorWIZARD_MACHINING_3
            
                If mbln_AdvanceMachiningOptions_3 Then .WizardStage = adoorWIZARD_LEADS
            
            Case adoorWIZARD_LEADS
            
                If mbln_AdvanceLeadInfo Then .WizardStage = adoorWIZARD_PATH_COMPLETE
                    
            Case adoorWIZARD_POCKET_LEADS
            
                If mbln_AdvancePocketLeads Then .WizardStage = adoorWIZARD_PATH_COMPLETE
            
            Case adoorWIZARD_MACHINING_STYLES
            
                If mbln_AdvanceMachiningStyles Then .WizardStage = adoorWIZARD_PATH_COMPLETE
                
            Case adoorWIZARD_PATH_COMPLETE
                                                    
                If mbln_AdvancePathComplete Then
                
                    .WizardStage = adoorWIZARD_CREATION_METHOD
    
                    '..add another path
                    .AddAnotherPath = True
    
                    '..is after the first path
                    .AfterFirstPath = True
          
                End If
          
        End Select
           
    End With
            
Controlled_Exit:

Exit Sub

m_WizardNext_Error:

    With App.Frame
        MsgBox Err.Description, vbExclamation, .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 50)
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "m_WizardNext"
    Resume Controlled_Exit

End Sub

Private Sub m_WizardBack()

On Error GoTo m_WizardBack_Error

    With clsWizard
    
        '..find the select wizard stage
        Select Case .WizardStage
        
            Case adoorWIZARD_DOOR_TYPE
                        
                '..should not happen
            
            Case adoorWIZARD_CREATION_METHOD
            
                .WizardStage = adoorWIZARD_DOOR_TYPE
            
            Case adoorWIZARD_MACHINE_METHOD
                        
                .WizardStage = adoorWIZARD_CREATION_METHOD
            
            Case adoorWIZARD_INSERT_PATH
                
                .WizardStage = adoorWIZARD_MACHINE_METHOD
                
            Case adoorWIZARD_OFFSET_AMOUNT
            
                If clsPathData.CreationMethod = DEF_CREATION_METHOD_MANUAL Then
                    .WizardStage = adoorWIZARD_MACHINE_METHOD
                Else
                    .WizardStage = adoorWIZARD_CREATION_METHOD
                End If
            
            Case adoorWIZARD_TOOL_DIRECTION_SIDE
                
                .WizardStage = adoorWIZARD_OFFSET_AMOUNT
            
            Case adoorWIZARD_POCKETING
            
                .WizardStage = adoorWIZARD_TOOL_DIRECTION_SIDE
            
            Case adoorWIZARD_ROUGH_FINISH
            
                .WizardStage = adoorWIZARD_TOOL_DIRECTION_SIDE
            
            Case adoorWIZARD_MACHINING_1
                
                If clsPathData.MachineMethod = DEF_MACHINE_METHOD_POCKET Then
                    
                    .WizardStage = adoorWIZARD_POCKETING
                    
                ElseIf clsPathData.MachineMethod = DEF_MACHINE_METHOD_ROUGHFINISH _
                  Or clsPathData.MachineMethod = DEF_MACHINE_METHOD_ENGRAVE Then
                
                    .WizardStage = adoorWIZARD_ROUGH_FINISH
                                    
                Else
                
                    .WizardStage = adoorWIZARD_TOOL_DIRECTION_SIDE
                                    
                End If
            
            Case adoorWIZARD_MACHINING_2
                        
                .WizardStage = adoorWIZARD_MACHINING_1
            
            Case adoorWIZARD_MACHINING_3
            
                .WizardStage = adoorWIZARD_MACHINING_2
            
            Case adoorWIZARD_LEADS
            
                If (clsPathData.MachineMethod = DEF_MACHINE_METHOD_ENGRAVE) Or (clsPathData.MachineMethod = DEF_MACHINE_METHOD_SIMPLE_ENGRAVE) Then
                
                    .WizardStage = adoorWIZARD_MACHINING_2
                
                ElseIf clsPathData.MachineMethod = DEF_MACHINE_METHOD_ROUGHFINISH Then
                
                    .WizardStage = adoorWIZARD_MACHINING_3
    
                End If
            
            Case adoorWIZARD_POCKET_LEADS
                
                .WizardStage = adoorWIZARD_MACHINING_2
            
            Case adoorWIZARD_MACHINING_STYLES
                        
                .WizardStage = adoorWIZARD_TOOL_DIRECTION_SIDE
            
            Case adoorWIZARD_PATH_COMPLETE
                
                '..this will be the FINISH button for this stage
                
                DoEvents
            
                Me.Hide
                
                '..no more paths
                .AddAnotherPath = False
                
                '..save the current door type name to the registry for the vb side to grab
                SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_TYPE, DEF_KEY_ACTIVE_TYPE, clsTypeData.TypeName
                
                '..looks like we made it so let the dll know
                SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, DEF_REG_KEY_CANCELWIZARD, 0
                                                                    
        End Select
            
    End With

Controlled_Exit:

Exit Sub

m_WizardBack_Error:

    With App.Frame
        MsgBox Err.Description, vbExclamation, .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 50)
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "m_WizardBack"
    Resume Controlled_Exit

End Sub

Private Function mbln_AdvanceDoorType() As Boolean
                                                
    Dim rstNew          As ADODB.Recordset
    Dim Fr              As Frame
    Dim sCTX            As String
    
On Error GoTo Error
    
    Set Fr = App.Frame
    
    sCTX = clsOptions.CTXFile
    
    '..ok so far
    mbln_AdvanceDoorType = True
                                                
    '..make sure there's a new type name
    If Len(Trim$(txtNewTypeName.Text)) = 0 Then
        
        '..let the user know what's going on
        MsgBox Fr.ReadTextFile(sCTX, 500, 51) & Space(3), vbInformation, DEF_PROJECT_NAME
        
        '..set focus to the new name
        txtNewTypeName.SetFocus
        
        '..no good yet
        mbln_AdvanceDoorType = False
        
        GoTo Controlled_Exit
        
    Else
        
        If Not clsTypeData.AlreadySet Then                                          '..07.18.02 - rg
        
            '..check for duplicate
            Set rstNew = grst_GetDoorTypeID(gstr_RemoveIllegalChars(txtNewTypeName.Text))
            
            If Not (rstNew Is Nothing) Then
                
                MsgBox Fr.ReadTextFile(sCTX, 500, 52) & Space(3), vbInformation, DEF_PROJECT_NAME
                       
                txtNewTypeName.Text = vbNullString
                txtNewTypeName.SetFocus
    
                '..no good here either
                mbln_AdvanceDoorType = False
                
                GoTo Controlled_Exit
            
            End If
        
        End If
        
    End If
    
    '..assign the user info
    With clsTypeData
        
        .TypeName = gstr_RemoveIllegalChars(txtNewTypeName.Text)
        .TypeDate = Now
        
        If Len(Trim$(txtNewTypeCreatedBy.Text)) = 0 Then
            .TypeCreatedBy = "UNKNOWN"
        Else
            .TypeCreatedBy = gstr_RemoveIllegalChars(txtNewTypeCreatedBy.Text)
        End If
        
        .TypeComments = gstr_RemoveIllegalChars(txtComments.Text)
    
    End With
                        
    '..insert the new type data                                             '..07.03.02 - rg
    If Not mbln_InsertTypeData() Then
        
        mbln_AdvanceDoorType = False
        GoTo Controlled_Exit
        
    End If
                        
    '..fill in some path data
    clsPathData.OperationNumber = 0
    
    '..is the first path
    clsWizard.AfterFirstPath = False
    
Controlled_Exit:
    
    Set rstNew = Nothing
    Set Fr = Nothing
    
Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceDoorType"
    mbln_AdvanceDoorType = False
    Call g_UnLoadAllForms
    
End Function

Private Function mbln_AdvanceMachiningMethod(iWiz As AdoorWizardStage) As Boolean
    
On Error GoTo Error
    
    '..ok so far
    mbln_AdvanceMachiningMethod = True
    
    With clsPathData
    
        '..set the machining method
        .MachineMethod = Switch(optMachiningMethodRoughFinish, DEF_MACHINE_METHOD_ROUGHFINISH, _
                                           optMachiningMethodPocket, DEF_MACHINE_METHOD_POCKET, _
                                           optMachiningMethodEngrave, DEF_MACHINE_METHOD_ENGRAVE, _
                                           optMachineMethodInsert, DEF_MACHINE_METHOD_INSERT, _
                                           optMachiningMethodSimpleEngrave, DEF_MACHINE_METHOD_SIMPLE_ENGRAVE)
        
        '..now lauch the next step depending on machine method
        Select Case .MachineMethod
    
            '..are we inserting?
            Case DEF_MACHINE_METHOD_INSERT: iWiz = adoorWIZARD_INSERT_PATH
                                        
            '..are we doing anything but inserting?
            Case Else: iWiz = adoorWIZARD_OFFSET_AMOUNT
                
        End Select
    
    End With
    
Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceMachiningMethod"
    mbln_AdvanceMachiningMethod = False
    Call g_UnLoadAllForms

End Function

Private Function mbln_AdvanceMachiningStyles() As Boolean
    
    Dim sMachiningStyle As String
    
On Error GoTo Error
    
    '..ok so far
    mbln_AdvanceMachiningStyles = True
    
    
    If TreeStyles.SelectedItem Is Nothing Then
        MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 73, 3) & Space(3), vbExclamation, DEF_PROJECT_NAME
        mbln_AdvanceMachiningStyles = False
        Exit Function
    End If
    
    If TreeStyles.SelectedItem.DataKey = "Folder" Then
        MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 73, 3) & Space(3), vbExclamation, DEF_PROJECT_NAME
        mbln_AdvanceMachiningStyles = False
        Exit Function
    End If
    
    sMachiningStyle = TreeStyles.SelectedItem.Tag
    
    With clsPathData
    
        .MachiningStyle = sMachiningStyle
        .MachineMethod = DEF_CREATION_METHOD_MACHINING_STYLE
        .FinalDepth = gdbl_Rounding(MillMachiningStyles(sMachiningStyle).MillSubStyles(1).GetMillData.FinalDepth, 3)
        .ToolName = MillMachiningStyles(sMachiningStyle).MillSubStyles(1).Tool.Name
        .ToolFullPath = gstr_StripLicomDatPath(MillMachiningStyles(sMachiningStyle).MillSubStyles(1).Tool.FileName)
    
        '..full or partial cut?
        .CutType = Switch(optStylesCutTypeFull.Value, DEF_CUT_TYPE_FULL, _
                         optStylesCutTypePartial.Value, DEF_CUT_TYPE_PARTIAL)
    
    End With
    
    ' Select Geometry
    Me.Hide
    
Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceMachiningStyles"
    mbln_AdvanceMachiningStyles = False
    Call g_UnLoadAllForms

End Function


Private Function mbln_AdvanceInsertPath() As Boolean
    
    Dim Ctl             As Control
    Dim Drw             As Drawing
    Dim MinX            As Double
    Dim MinY            As Double
    Dim MaxX            As Double
    Dim MaxY            As Double
    Dim sNewName        As String
    Dim pOut            As Path
    
On Error GoTo Error
    
    '..start out ok
    mbln_AdvanceInsertPath = True
    
    '..do this to force the BeforeUpdate event to all the text boxes
    cmdNext.SetFocus
    
    DoEvents
    
    '..make sure a file has been selected
    If Len(lblInsertFilePath.Caption) = 0 Then
    
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 53) & Space(3), vbInformation, DEF_PROJECT_NAME
        mbln_AdvanceInsertPath = False
        GoTo Controlled_Exit
        
    End If
    
    '..make sure all data is entered
    If Not mbln_CheckAllText(Ctl) Then
    
        mbln_AdvanceInsertPath = False
        GoTo Controlled_Exit
        
    End If
    
    Set Drw = App.ActiveDrawing
    
    '..fill in the path data
    With clsPathData
        
        sNewName = gstr_ReplaceApostrophe(Trim$(lblInsertFilePath.Caption))
        
        '..rename the drawing
        Name Trim$(lblInsertFilePath.Caption) As sNewName
        
        '..fill in the insert parameters
        .InsertFilePath = sNewName                      ' Trim$(lblInsertFilePath.Caption)
        .InsertFilePointX = PDbl(txtDistanceFromRefX)
        .InsertFilePointY = PDbl(txtDistanceFromRefY)
        .InsertFileGroupNumber = PDbl(txtParametricGroupNumber)
        
        '..are we aut-centering X?
        If chkInsertFileCenterX.Value Then
        
            If chkInsertFileCenterY.Value = True Then
                .InsertFileReferencePoint = adoorINSERT_AUTO_CENTER
            Else
                .InsertFileReferencePoint = IIf(optRefTop.Value, adoorINSERT_CENTER_TOP_X, adoorINSERT_CENTER_BOTTOM_X)
            End If
                    
        '..are we centering Y?
        ElseIf chkInsertFileCenterY.Value Then
            
            If chkInsertFileCenterX.Value = True Then
                .InsertFileReferencePoint = adoorINSERT_AUTO_CENTER
            Else
                .InsertFileReferencePoint = IIf(optRefLeft.Value, adoorINSERT_CENTER_LEFT_Y, adoorINSERT_CENTER_RIGHT_Y)
            End If
        
        Else
        
            .InsertFileReferencePoint = Switch(optRefTopLeft.Value, adoorINSERT_TOP_LEFT, _
                                               optRefTopRight.Value, adoorINSERT_TOP_RIGHT, _
                                               optRefBottomLeft.Value, adoorINSERT_BOTTOM_LEFT, _
                                               optRefBottomRight.Value, adoorINSERT_BOTTOM_RIGHT, _
                                               optRefParametric.Value, adoorINSERT_PARAMETRIC)
                                           
        End If
                                           
        '..set active tool as dummy tool
        .ToolFullPath = DEF_INSERT_DUMMY_TOOL
        .ToolName = DEF_INSERT_DUMMY_TOOL
        
        '..loop, find the outside geo and get the extents
        For Each pOut In Drw.Geometries
        
            If pOut.Attribute(DEF_ATT_GEOMETRY_NUMBER) = 1 Then
                
                pOut.GetFeedExtent MinX, MinY, MaxX, MaxY
                Exit For
                
            End If
        
        Next pOut
    
        '..find the reference point
        Select Case .InsertFileReferencePoint
            
            Case adoorINSERT_AUTO_CENTER
            
                Drw.InsertDrawing .InsertFilePath, MaxX / 2, MaxY / 2, 0
                
            Case adoorINSERT_CENTER_BOTTOM_X
            
                Drw.InsertDrawing .InsertFilePath, MaxX / 2, MinY + .InsertFilePointY, 0
                
            Case adoorINSERT_CENTER_TOP_X
            
                Drw.InsertDrawing .InsertFilePath, MaxX / 2, MaxY - .InsertFilePointY, 0
                
            Case adoorINSERT_CENTER_LEFT_Y
            
                Drw.InsertDrawing .InsertFilePath, MinX + .InsertFilePointX, MaxY / 2, 0
                
            Case adoorINSERT_CENTER_RIGHT_Y
            
                Drw.InsertDrawing .InsertFilePath, MaxX - .InsertFilePointX, MaxY / 2, 0
            
            Case adoorINSERT_TOP_LEFT
                
                Drw.InsertDrawing .InsertFilePath, MinX + .InsertFilePointX, MaxY - .InsertFilePointY, 0
                
            Case adoorINSERT_TOP_RIGHT
            
                Drw.InsertDrawing .InsertFilePath, MaxX - .InsertFilePointX, MaxY - .InsertFilePointY, 0
            
            Case adoorINSERT_BOTTOM_LEFT
            
                Drw.InsertDrawing .InsertFilePath, MinX + .InsertFilePointX, MinY + .InsertFilePointY, 0
            
            Case adoorINSERT_BOTTOM_RIGHT
            
                Drw.InsertDrawing .InsertFilePath, MaxX - .InsertFilePointX, MinY + .InsertFilePointY, 0
            
            Case adoorINSERT_PARAMETRIC
            
                If Not gbln_InsertDrawing(.InsertFilePath, .InsertFileGroupNumber) Then
                    mbln_AdvanceInsertPath = False
                    GoTo Controlled_Exit
                End If
            
        End Select
                    
    End With
    
    '..hide me to show the insert and add the data to db
    Me.Hide
    
Controlled_Exit:
    
    Set Drw = Nothing
    
Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceInsertPath"
    mbln_AdvanceInsertPath = False
    Call g_UnLoadAllForms
    
End Function

Private Sub chkInsertFileCenterX_Click() 'Value As Integer)

On Error Resume Next

    '..enable and disable controls as necessary
    If chkInsertFileCenterX.Value Then
        
        optRefBottomLeft.Enabled = False
        optRefBottomRight.Enabled = False
        optRefTopLeft.Enabled = False
        optRefTopRight.Enabled = False
        
        optRefBottomLeft.Visible = False
        optRefBottomRight.Visible = False
        optRefTopLeft.Visible = False
        optRefTopRight.Visible = False
        
        optRefRight.Visible = True
        optRefLeft.Visible = True
        optRefTop.Visible = True
        optRefBottom.Visible = True
        
        If chkInsertFileCenterY.Value = True Then
            
            fraInsertPositionPoint.Enabled = False
            lblInsertPointX.Enabled = False
            lblInsertPointY.Enabled = False
            txtDistanceFromRefX.Enabled = False
            txtDistanceFromRefY.Enabled = False
            
            optRefRight.Enabled = False
            optRefLeft.Enabled = False
            optRefTop.Enabled = False
            optRefBottom.Enabled = False
                            
        Else
        
            fraInsertPositionPoint.Enabled = True
            lblInsertPointX.Enabled = False
            lblInsertPointY.Enabled = True
            txtDistanceFromRefX.Enabled = False
            txtDistanceFromRefY.Enabled = True
            
            optRefRight.Enabled = False
            optRefLeft.Enabled = False
            optRefTop.Enabled = True
            optRefBottom.Enabled = True
            
            optRefTop.Value = True
                            
        End If
        
    Else
        
        If chkInsertFileCenterY.Value = True Then
            
            optRefBottomLeft.Enabled = False
            optRefBottomRight.Enabled = False
            optRefTopLeft.Enabled = False
            optRefTopRight.Enabled = False
            
            optRefBottomLeft.Visible = False
            optRefBottomRight.Visible = False
            optRefTopLeft.Visible = False
            optRefTopRight.Visible = False
            
            optRefRight.Visible = True
            optRefLeft.Visible = True
            optRefTop.Visible = True
            optRefBottom.Visible = True
            
            optRefRight.Enabled = True
            optRefLeft.Enabled = True
            optRefTop.Enabled = False
            optRefBottom.Enabled = False
            
            optRefLeft.Value = True
            
            fraInsertPositionPoint.Enabled = True
            lblInsertPointX.Enabled = True
            lblInsertPointY.Enabled = False
            txtDistanceFromRefX.Enabled = True
            txtDistanceFromRefY.Enabled = False
            
        Else
                    
            optRefBottomLeft.Visible = True
            optRefBottomRight.Visible = True
            optRefTopLeft.Visible = True
            optRefTopRight.Visible = True
                        
            optRefBottomLeft.Enabled = True
            optRefBottomRight.Enabled = True
            optRefTopLeft.Enabled = True
            optRefTopRight.Enabled = True
            
            optRefTopLeft.Value = True

            optRefRight.Enabled = False
            optRefLeft.Enabled = False
            optRefTop.Enabled = False
            optRefBottom.Enabled = False
            
            optRefRight.Visible = False
            optRefLeft.Visible = False
            optRefTop.Visible = False
            optRefBottom.Visible = False
            
            fraInsertPositionPoint.Enabled = True
            lblInsertPointX.Enabled = True
            lblInsertPointY.Enabled = True
            txtDistanceFromRefX.Enabled = True
            txtDistanceFromRefY.Enabled = True
            
        End If

    End If

    If (Err.Number <> 0) Then WriteError Err, False, "chkInsertFileCenterX_Click"
    
End Sub

Private Sub cmdBrowse_Click()
    
    Dim Drw             As Drawing
    Dim drwTemp         As Drawing
    Dim MinX            As Double
    Dim MinY            As Double
    Dim minZ            As Double
    Dim MaxX            As Double
    Dim MaxY            As Double
    Dim maxZ            As Double
    Dim minXtmp         As Double
    Dim minYtmp         As Double
    Dim minZtmp         As Double
    Dim maxXtmp         As Double
    Dim maxYtmp         As Double
    Dim maxZtmp         As Double
    Dim Fr              As Frame
    Dim cDialog         As New CFileDialog
    Dim sCTX            As String
    Dim sTmp            As String
    Dim FSO             As New Scripting.FileSystemObject
    
On Error GoTo cmdBrowse_Error

    Set Fr = App.Frame
    
    sTmp = gs_GetCommonAppDataDir & "tmp.ard"
                      
    Set Drw = App.ActiveDrawing
    
    g_ConvertGroupsToAttributes
    Drw.SaveAs sTmp

    sCTX = clsOptions.CTXFile
    
    With cDialog
        
        '..clear the file name
        .DefaultExt = "ard"
        .DialogTitle = Fr.ReadTextFile(sCTX, 10, 18)
        .Filter = "AlphaCAM Router Drawing (.ard)|*.ard"
        .FilterIndex = 0
        .hWndParent = FindWindow32(vbNullString, App.Name)
        .FileName = ""
        .InitialDir = gstr_CheckDir(App.LicomdirPath) & DEF_LICOMDIR
        
        '..if we have something?
        If .Show(True) Then
                                                    
            '..get the extents of the current drawing
            Call Drw.GetExtent(MinX, MinY, minZ, MaxX, MaxY, maxZ)
                            
            '..let's make sure there are toolpaths in the selected file
            Set drwTemp = App.OpenDrawing(.FileName)
            
            If drwTemp.GetToolPathCount = 0 Then
                
                MsgBox Fr.ReadTextFile(sCTX, 500, 54) & Space(3) & vbCrLf & _
                       Fr.ReadTextFile(sCTX, 500, 55) & Space(3) & vbCrLf & _
                       vbCrLf & _
                       Fr.ReadTextFile(sCTX, 500, 57) & Space(3), vbInformation, DEF_PROJECT_NAME
                       
                GoTo Controlled_Exit
                
            End If
                        
            '..look for workplanes
            If (drwTemp.WorkPlanes.Count <> 0) Then
                MsgBox Fr.ReadTextFile(sCTX, 600, 146) & Space(3), vbInformation, DEF_PROJECT_NAME
                GoTo Controlled_Exit
            End If
            
            '..get the extents of the file to be inserted
            Call drwTemp.GetExtent(minXtmp, minYtmp, minZtmp, maxXtmp, maxYtmp, maxZtmp)
            
            '..check insert dimensions against current door dimensions to make sure it will fit
            If ((maxXtmp - minXtmp) > (MaxX - MinX)) Then
                            
                '..it's too big, ask what to do
                If MsgBox(Fr.ReadTextFile(sCTX, 500, 58) & Space(3) & vbCrLf & _
                          Fr.ReadTextFile(sCTX, 500, 59) & Space(3), _
                          vbQuestion + vbYesNo, DEF_PROJECT_NAME) = vbNo Then GoTo Controlled_Exit
            
            ElseIf ((maxYtmp - minYtmp) > (MaxY - MinY)) Then
                
                '..it's too big, ask what to do
                If MsgBox(Fr.ReadTextFile(sCTX, 500, 58) & Space(3) & vbCrLf & _
                          Fr.ReadTextFile(sCTX, 500, 59) & Space(3), _
                          vbQuestion + vbYesNo, DEF_PROJECT_NAME) = vbNo Then GoTo Controlled_Exit
                
            End If
                            
            lblInsertFilePath.Caption = Space(1) & .FileName

        End If
        
    End With

Controlled_Exit:
    
    '..reload the temp drawing
    If FSO.FileExists(sTmp) Then App.OpenDrawing sTmp
    g_ConvertAttributesToGroups
    
    Set Drw = Nothing
    Set drwTemp = Nothing
    Set Fr = Nothing
    Set cDialog = Nothing
    Set FSO = Nothing

Exit Sub

cmdBrowse_Error:

    MsgBox Fr.ReadTextFile(sCTX, 500, 60) & Space(3), vbExclamation, DEF_PROJECT_NAME
    If (Err.Number <> 0) Then WriteError Err, True, "frmWizard_cmdBrowse_Click"
    Resume Controlled_Exit

End Sub

Private Sub txtDistanceFromRefX_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    '..eval value
    gbln_TextCalc txtDistanceFromRefX, Cancel
    
End Sub

Private Sub txtDistanceFromRefY_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    '..eval value
    gbln_TextCalc txtDistanceFromRefY, Cancel
    
End Sub

Private Function mbln_AdvanceOffsetAmount() As Boolean
    
    Dim Ctl             As Control
    
On Error GoTo Error
    
    '..start out fine
    mbln_AdvanceOffsetAmount = True
    
    '..do this to force the BeforeUpdate event to all the text boxes
    cmdNext.SetFocus
    
    DoEvents
    
    '..make sure all data is entered
    If Not mbln_CheckAllText(Ctl) Then
    
        mbln_AdvanceOffsetAmount = False
        
        GoTo Controlled_Exit
        
    End If
    
    '..check to see if they've select a side with no offset
    If Not optOffsetSideCenter Then
        
        '..is the offset value greater then zero?
        If PDbl(txtOffset) = 0 Then
            
            '..no, so ask for a value
            MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 61) & Space(3), vbInformation, DEF_PROJECT_NAME
            
            '..focus on the needed value
            txtOffset.SetFocus
            
            mbln_AdvanceOffsetAmount = False
            
            '..don't go anywhere
            GoTo Controlled_Exit
            
        End If
        
    End If
    
    With clsPathData
        
        '..set the offset value
        .PathOffsetValue = PDbl(txtOffset)
    
        '..get the offset side
        .PathOffsetSide = Switch(optOffsetSideInside, acamINSIDE, _
                             optOffsetSideCenter, acamCENTER, _
                             optOffsetSideOutside, acamOUTSIDE)
                        
        .LeadEntryPointIsCorner = Switch(optLeadEntryPointCorner, AdoorLeadEntryPoint_Corner, _
                             optLeadEntryPointMiddle, AdoorLeadEntryPoint_Midpoint, _
                             optLeadEntryPointDrawn, AdoorLeadEntryPoint_Drawn)

        '..make sure no insert file path gets passed from possible previous
        .InsertFilePath = "NONE"
        .InsertFilePointX = 0
        .InsertFilePointY = 0
        .InsertFileReferencePoint = 0
           
    End With
    
Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceOffsetAmount"
    mbln_AdvanceOffsetAmount = False
    Resume Controlled_Exit
    
End Function

Private Function mbln_AdvanceToolDirectionSide(iWiz As AdoorWizardStage) As Boolean
    
    Dim blnCancelToolSelect     As Boolean
    
On Error GoTo Error
    
    '..start out ok
    mbln_AdvanceToolDirectionSide = True
    
    '..set the mill data
    If mbln_SetToolSideData Then
        
        With clsPathData
        
            If .CreationMethod = DEF_CREATION_METHOD_MACHINING_STYLE Then
            
                iWiz = adoorWIZARD_MACHINING_STYLES
            
            Else
            
                '..get the selected machining method and launch the options
                If .MachineMethod = DEF_MACHINE_METHOD_ROUGHFINISH Then
                    
                    '..ask for a tool
                    iWiz = IIf(gbln_PickTool, adoorWIZARD_ROUGH_FINISH, clsWizard.WizardStage)
            
                ElseIf .MachineMethod = DEF_MACHINE_METHOD_POCKET Then
                
                    '..ask for a tool
                    iWiz = IIf(gbln_PickTool, adoorWIZARD_POCKETING, clsWizard.WizardStage)
                
                ElseIf (.MachineMethod = DEF_MACHINE_METHOD_ENGRAVE) Or (.MachineMethod = DEF_MACHINE_METHOD_SIMPLE_ENGRAVE) Then
                        
                    '..ask for a tool
                    If gbln_PickTool Then
                        
                        '..ask for a conical tool until we get one
                        Do While Not App.GetCurrentTool.Type = acamToolUSER
                            
                            MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 62) & Space(3), vbExclamation, DEF_PROJECT_NAME
                                       
                            If gbln_PickTool Then
                                
                                '..should have picked one
                                blnCancelToolSelect = False
                            
                            Else
                                                             
                                '..canceled tool selection
                                blnCancelToolSelect = True
                                
                                '..get out of the loop
                                Exit Do
    
                            End If
                            
                        Loop
    
                        '..go to the next step or stay put?
                        iWiz = IIf(blnCancelToolSelect, clsWizard.WizardStage, adoorWIZARD_ROUGH_FINISH)
    
                        '..make sure the is no pocket Boundary distance
                        .PocketBoundary = 0
                        
                    Else
    
                        '..no tool picked so stay put
                        iWiz = clsWizard.WizardStage
                        
                        GoTo Controlled_Exit
            
                    End If
                
                End If
        
            End If
        
        End With
        
        
    Else
        
        '..let the user know there's a problem
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 63) & Space(3), vbExclamation, DEF_PROJECT_NAME
        mbln_AdvanceToolDirectionSide = False
        GoTo Controlled_Exit
        
    End If
    
Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceToolDirectionSide"
    mbln_AdvanceToolDirectionSide = False
    Call g_UnLoadAllForms

End Function

Private Function mbln_AdvancePathComplete() As Boolean
    
On Error GoTo cmd_Error
    
    '..start out ok
    mbln_AdvancePathComplete = True
        
    With clsWizard
    
        '..add another path
        .AddAnotherPath = True
        
        '..is after the first path
        .AfterFirstPath = True
        
    End With

Controlled_Exit:

Exit Function

cmd_Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvancePathComplete"
    mbln_AdvancePathComplete = False
    Call g_UnLoadAllForms

End Function

Private Function mbln_AdvanceMachiningOptions_1() As Boolean
            
    Dim Ctl         As Control
    
On Error GoTo Error
    
    '..do this to force the BeforeUpdate event to all the text boxes
    cmdNext.SetFocus
    
    '..start out ok
    mbln_AdvanceMachiningOptions_1 = True
    
    '..make sure all data is entered
    If Not mbln_CheckAllText(Ctl) Then
        
        mbln_AdvanceMachiningOptions_1 = False
        GoTo Controlled_Exit
        
    End If
        
    If mbln_SetMachiningOptionsData_1 Then
        
        With clsPathData
        
            '..check multiple depths
            If .MultiplePasses Then
                        
                If Not mbln_CheckMultipleDepths Then
                    
                    '..set the focus to the first cut depth and select it
                    With txtThicknessOfFirstCut
                        .SetFocus
                        .SelStart = 0
                        .SelLength = Len(.Text)
                    End With
                    
                    mbln_AdvanceMachiningOptions_1 = False
                    GoTo Controlled_Exit
                
                End If
            
            End If
        
            '..look for no depth or depth above material when engraving
            If (.MachineMethod = DEF_MACHINE_METHOD_ENGRAVE) Or (.MachineMethod = DEF_MACHINE_METHOD_SIMPLE_ENGRAVE) Then
                
                If .IsFinalDepthPercent Then
                    
                    If (.FinalDepthPercentage + .MaterialTop) = 0 Then

                        With txtFinalDepthPercent
                            .SetFocus
                            .SelStart = 0
                            .SelLength = Len(.Text)
                        End With
                        
                        Beep
                        mbln_AdvanceMachiningOptions_1 = False
                        GoTo Controlled_Exit
                                            
                    ElseIf .FinalDepthPercentage < 0 Then
                                                
                        With txtFinalDepthPercent
                            .SetFocus
                            .SelStart = 0
                            .SelLength = Len(.Text)
                        End With
                        
                        Beep
                        mbln_AdvanceMachiningOptions_1 = False
                        GoTo Controlled_Exit
                        
                    End If
                
                Else
                
                    If (.FinalDepth >= .MaterialTop) Then
                        
                        With txtFinalDepth
                            .SetFocus
                            .SelStart = 0
                            .SelLength = Len(.Text)
                        End With
                        
                        Beep
                        mbln_AdvanceMachiningOptions_1 = False
                        GoTo Controlled_Exit
                        
                    End If
                
                End If
                
            End If
        
        End With
        
    Else
        
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 64) & Space(3), vbExclamation, DEF_PROJECT_NAME
        mbln_AdvanceMachiningOptions_1 = False
        GoTo Controlled_Exit
        
    End If

Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceMachiningOptions_1"
    mbln_AdvanceMachiningOptions_1 = False
    Call g_UnLoadAllForms

End Function

Private Function mbln_AdvanceMachiningOptions_2(iWiz As AdoorWizardStage) As Boolean

    Dim Ctl         As Control
    
On Error GoTo Error
    
    '..do this to force the BeforeUpdate event to all the text boxes
    cmdNext.SetFocus
    
    '..start out ok
    mbln_AdvanceMachiningOptions_2 = True
    
    '..make sure all data is entered
    If Not mbln_CheckAllText(Ctl) Then
    
        mbln_AdvanceMachiningOptions_2 = False
        GoTo Controlled_Exit
        
    End If
        
    With txtToolNumber
        
        '..we need a tool number (we multiply by 1 in case there's more than one 0 in the field)
        If PDbl(.Text) * 1 = 0 Then
            
            Beep
            
            mbln_AdvanceMachiningOptions_2 = False
            
            .SetFocus
            .SelStart = 0
            .SelLength = Len(txtToolNumber)
            
            GoTo Controlled_Exit
        
        End If
        
    End With
        
    '..try to set the mill data
    If mbln_SetMachiningOptionsData_2 Then
        
        Select Case clsPathData.MachineMethod
        
            Case DEF_MACHINE_METHOD_ROUGHFINISH
        
                iWiz = adoorWIZARD_MACHINING_3
        
            Case DEF_MACHINE_METHOD_ENGRAVE, DEF_MACHINE_METHOD_SIMPLE_ENGRAVE
                
                iWiz = adoorWIZARD_LEADS
            
            Case DEF_MACHINE_METHOD_POCKET
            
                iWiz = adoorWIZARD_POCKET_LEADS
            
            Case Else
            
                DoEvents
    
                '..hide me to pick geometry
                Me.Hide
            
                '..disable the buttons (trust me)
                cmdNext.Enabled = False
                cmdCancel.Enabled = False
                cmdBack.Enabled = False
                
                iWiz = adoorWIZARD_PATH_COMPLETE
                
        End Select
    
    Else
        
        '..let the user know there's a problem
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 64) & Space(3), vbExclamation, DEF_PROJECT_NAME
        mbln_AdvanceMachiningOptions_2 = False
        GoTo Controlled_Exit
        
    End If
    
Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceMachiningOptions_2"
    mbln_AdvanceMachiningOptions_2 = False
    Call g_UnLoadAllForms
    
End Function

Private Function mbln_AdvanceMachiningOptions_3() As Boolean

    Dim Ctl         As Control
    
On Error GoTo Error
    
    '..do this to force the BeforeUpdate event to all the text boxes
    cmdNext.SetFocus
    
    '..start out ok
    mbln_AdvanceMachiningOptions_3 = True
    
    '..make sure all data is entered
    If Not mbln_CheckAllText(Ctl) Then
    
        mbln_AdvanceMachiningOptions_3 = False
        GoTo Controlled_Exit
        
    End If
        
    '..try to set the mill data
    If Not mbln_SetMachiningOptionsData_3 Then
        
        '..let the user know there's a problem
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 64) & Space(3), vbExclamation, DEF_PROJECT_NAME
        mbln_AdvanceMachiningOptions_3 = False
        GoTo Controlled_Exit
        
    End If
    
Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceMachiningOptions_2"
    mbln_AdvanceMachiningOptions_3 = False
    Call g_UnLoadAllForms
    
End Function


Private Function mbln_AdvanceLeadInfo() As Boolean

    Dim Ctl         As Control
    
On Error GoTo cmd_Error
    
    '..do this to force the BeforeUpdate event to all the text boxes
    cmdNext.SetFocus
    
    '..start out ok
    mbln_AdvanceLeadInfo = True
    
    '..make sure all data is entered
    If Not mbln_CheckAllText(Ctl) Then
    
        mbln_AdvanceLeadInfo = False
        GoTo Controlled_Exit
        
    End If
    
    If mbln_SetLeadData Then
    
        ' 11/21/05 - rg
        '
        '..check comp method
        If (clsPathData.MachineCompensation <> acamCompTOOLCEN) And (StrComp(clsPathData.MachineMethod, DEF_MACHINE_METHOD_ROUGHFINISH) = 0) Then
            
            '..check lead method
            If optLeadInArc Or optLeadOutArc Or _
               optLeadInNone Or optLeadOutNone Then
                
                With App.Frame
                
                    '..ask the user to continue
                    If MsgBox(.ReadTextFile(clsOptions.CTXFile, 500, 65) & Space(3) & vbCrLf & _
                              .ReadTextFile(clsOptions.CTXFile, 500, 66) & Space(3) & vbCrLf & _
                              .ReadTextFile(clsOptions.CTXFile, 500, 67) & Space(3), _
                              vbExclamation + vbOKCancel, DEF_PROJECT_NAME) = vbCancel Then
                           
                           mbln_AdvanceLeadInfo = False
                           
                           GoTo Controlled_Exit
                    
                    End If
                
                End With
                
            End If
                
        End If
                
        ' Test for toolside set to centre (manual lead-in/out)
        If (clsPathData.ToolSide = acamCENTER) Or _
           (StrComp(clsPathData.MachineMethod, DEF_MACHINE_METHOD_ENGRAVE) = 0) Or _
           (StrComp(clsPathData.MachineMethod, DEF_MACHINE_METHOD_SIMPLE_ENGRAVE) = 0) Then
            
            If Not (optLeadInNone.Value And optLeadOutNone.Value) Then
            
                If Not ((optLeadInLine.Value And chkLeadInSloping.Value = True) And _
                  (optLeadOutLine.Value And chkLeadOutSloping.Value = True)) Or _
                  PDbl(txtLeadApproachAngle) <> 0 Then
                  
                  With Frame
                    MsgBox .ReadTextFile(clsOptions.CTXFile, 600, 184) & Chr(13) & _
                      .ReadTextFile(clsOptions.CTXFile, 600, 185), vbExclamation, DEF_PROJECT_NAME
                  End With
                  
                  mbln_AdvanceLeadInfo = False
                  GoTo Controlled_Exit
                  
                End If
              
            End If
        End If
        
        
        If optLeadInLine And Val(txtLeadLineLengthIn) = 0 Then
            MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 112) & Space(3), vbExclamation, DEF_PROJECT_NAME
            
            mbln_AdvanceLeadInfo = False
            
            GoTo Controlled_Exit
        End If
        
        If optLeadOutLine And Val(txtLeadLineLengthOut) = 0 Then
            MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 113) & Space(3), vbExclamation, DEF_PROJECT_NAME
            
            mbln_AdvanceLeadInfo = False
            
            GoTo Controlled_Exit
        End If
        
        If (optLeadInArc Or optLeadOutArc) And Val(txtLeadArcRadius) = 0 Then
            MsgBox Frame.ReadTextFile(clsOptions.CTXFile, 500, 114) & Space(3), vbExclamation, DEF_PROJECT_NAME
            
            mbln_AdvanceLeadInfo = False
            
            GoTo Controlled_Exit
        End If
        
        DoEvents
    
        '..hide me and select geometry
        Me.Hide
    
    Else
        
        '..let the user know there's a problem
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 64) & Space(3), vbExclamation, DEF_PROJECT_NAME
        mbln_AdvanceLeadInfo = False
        GoTo Controlled_Exit
        
    End If

Controlled_Exit:

Exit Function

cmd_Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceLeadInfo"
    mbln_AdvanceLeadInfo = False
    Call g_UnLoadAllForms
    
End Function

Private Function mbln_AdvancePocket() As Boolean
    
On Error GoTo Error
    
    '..start out groovy
    mbln_AdvancePocket = True
    
    cmdNext.SetFocus
    
    With txtPocketBoundary
        
        '..see if there is going to be a Boundary
        If (PDbl(.Text) <> 0) Then
            
            '..let's make sure the Boundary value is greater than the tool diameter
            If (PDbl(.Text) <= App.GetCurrentTool.Diameter) Then
                
                '..let the user know what the hell is going on
                MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 68) & Space(3), vbInformation, DEF_PROJECT_NAME
                
                '..focus in on the desired value
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
                        
                '..bolt
                mbln_AdvancePocket = False
                
                GoTo Controlled_Exit
                
            End If
        
        End If
        
    End With
    
    '..set the mill data
    If Not mbln_SetPocketData Then
        
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 69) & Space(3), vbExclamation, DEF_PROJECT_NAME
        mbln_AdvancePocket = False
        GoTo Controlled_Exit
    
    End If

Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvancePocket"
    mbln_AdvancePocket = False
    Call g_UnLoadAllForms

End Function

Private Function mbln_AdvanceRoughFinish() As Boolean
    
On Error GoTo Error
    
    '..start out ok
    mbln_AdvanceRoughFinish = True
    
    If Not mbln_SetRoughFinishData Then
        
        '..let the user know there's a problem
        MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 64) & Space(3), vbExclamation, DEF_PROJECT_NAME
               
        mbln_AdvanceRoughFinish = False
        
    End If
    
Controlled_Exit:

Exit Function

Error:

    With App.Frame

        MsgBox Err.Description & vbCrLf & vbCrLf & _
               .ReadTextFile(clsOptions.CTXFile, 500, 44) & Space(1) & _
               .ReadTextFile(clsOptions.CTXFile, 500, 41) & Space(3), vbExclamation, DEF_PROJECT_NAME
    
    End With
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_AdvanceRoughFinish"
    mbln_AdvanceRoughFinish = False
    Call g_UnLoadAllForms

End Function

Private Sub optOffsetSideInside_Click()

    lblOffset.Enabled = True
    txtOffset.Enabled = True

    If (Err.Number <> 0) Then WriteError Err, False, "optOffsetSideInside_Click"
    
End Sub

Private Sub optOffsetSideOutside_Click()
    
    lblOffset.Enabled = True
    txtOffset.Enabled = True

    If (Err.Number <> 0) Then WriteError Err, False, "optOffsetSideOutside_Click"
    
End Sub

Private Sub txtOffset_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtOffset, Cancel
    
End Sub

Private Function mbln_InsertTypeData() As Boolean                                   '..07.03.02 - rg

    Dim rstType                 As ADODB.Recordset

    
On Error GoTo mbln_InsertTypeData_Error

    '..start our aweful
    mbln_InsertTypeData = False

    With App.Frame
    
        '..let the user know what's going on
        .SetProgressText .ReadTextFile(clsOptions.CTXFile, 300, 17) & Space(3)
        DoEvents
        
    End With
    
    '..connect to the database
    If Not gbln_ConnectToDB Then Exit Function
        
    With clsTypeData

        '..already been set?
        If .AlreadySet Then
            
            '..already been set to get the current data
            Set rstType = grst_GetDoorTypeData(.CurrentPK)
            
            If (rstType Is Nothing) Then mbln_InsertTypeData = False: GoTo Controlled_Exit
        
        Else
            
            '..not set yet so start new recordset
            Set rstType = New ADODB.Recordset                                       '..07.17.02 - rg
            
            rstType.CursorType = adOpenKeyset
            rstType.LockType = adLockOptimistic
            rstType.Open "AD_DOOR_TYPES", gdb_CDM, , , adCmdTable
    
            rstType.AddNew

        End If

        rstType.Fields!TypeID = .TypeName
        rstType.Fields!DateAdded = .TypeDate
        rstType.Fields!CreatorName = .TypeCreatedBy
        rstType.Fields!Comment = .TypeComments
        rstType.Fields!UserStyleName = .VBAMacroStyle
        rstType.Fields!Width = 0
        rstType.Fields!Length = 0
        rstType.Fields!CornerRadius = 0
        
        'iBypassNesting = CStr(GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "BypassNesting", 0))   ' removed 07/28/03 - rg
        'rstType.Fields!ByPassNest = iBypassNesting                                                         ' removed 07/28/03 - rg
        rstType.Fields!ByPassNest = False                                                                   ' reinstated 07/28/03 - rg
        
        rstType.Fields!RotationMethod = 0
        rstType.Fields!RotationAngle = 90
        rstType.Fields!UserStyle = True
        rstType.Fields!UserVariableString = vbNullString
        rstType.Fields!UserDescriptionString = vbNullString
        rstType.Fields!UserValue_0 = vbNullString
        rstType.Fields!UserValue_1 = vbNullString
        rstType.Fields!UserValue_2 = vbNullString
        rstType.Fields!UserValue_3 = vbNullString
        rstType.Fields!UserValue_4 = vbNullString
        rstType.Fields!UserValue_5 = vbNullString
        rstType.Fields!UserValue_6 = vbNullString
        
        rstType.Update
        
        '..set flag and get id just in case we come back here
        .AlreadySet = True
        .CurrentPK = rstType.Fields!PK
        
    End With
    
    mbln_InsertTypeData = True

Controlled_Exit:

    If Not (rstType Is Nothing) Then
        With rstType
            If (.State = adStateOpen) Then .Close
        End With
    End If

Exit Function

mbln_InsertTypeData_Error:
    
    MsgBox Err.Description, vbExclamation, "mbln_InsertTypeData"
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_InsertTypeData"
    mbln_InsertTypeData = False
    
    If Not (rstType Is Nothing) Then
        With rstType
            If (.EditMode = adEditAdd Or adEditInProgress) Then .CancelUpdate
        End With
    End If
    
    App.Frame.CloseProgressBox
    
    Resume Controlled_Exit
        
End Function

Private Function mbln_SetToolSideData() As Boolean
    
On Error GoTo mbln_SetMillData_Error
    
    '..start out cool
    mbln_SetToolSideData = True
    
    '..load in the mill data
    With clsPathData
        
        '..tool direction open
        .ToolDirectionCW = optDirectionCW
        
        '..tool direction closed
        .ToolDirectionIsReversed = optDirectionReversed
        
        '..inside/outside
        .ToolInOut = mint_CheckNullToolSide(Switch(optSideInside.Value, acamINSIDE, _
                                                   optSideOutside.Value, acamOUTSIDE, _
                                                   optSideCentreClosed.Value, acamON_CENTER))
                        
        '..left/right
        .ToolSide = mint_CheckNullToolSide(Switch(optSideRight.Value, acamRIGHT, _
                                                  optSideLeft.Value, acamLEFT, _
                                                  optSideCentreOpen.Value, acamCENTER))
        
        '..reverse partial paths
        .ToolSidePartialReverse = chkReverseToolDirection.Value
        
    End With
    
Exit Function
    
'..damnit
mbln_SetMillData_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_SetToolSideData"
    mbln_SetToolSideData = False

End Function

Private Function mbln_SetLeadData() As Boolean

On Error GoTo mbln_SetLeadData_Error

    '..start out fine
    mbln_SetLeadData = True

    With clsPathData
            
        '..fill up the lead info
        .LeadLineLengthIn = PDbl(txtLeadLineLengthIn.Text)                      '..07.18.02 - rg
        .LeadLineLengthOut = PDbl(txtLeadLineLengthOut.Text)                    '..07.18.02 - rg
        .LeadArcRadius = PDbl(txtLeadArcRadius.Text)
        .LeadApproachAngle = PDbl(txtLeadApproachAngle.Text)
        .LeadOverlap = PDbl(txtLeadOverlap.Text)
        
        .LeadInType = Switch(optLeadInLine.Value, acamLeadLINE, _
                             optLeadInArc.Value, acamLeadARC, _
                             (optLeadInBoth.Value And chkLeadLineArcTangentialIn.Value), acamLeadBOTH, _
                             (optLeadInBoth.Value And (Not chkLeadLineArcTangentialIn.Value)), acamLeadBOTH_NOT_TANGENTIAL, _
                             optLeadInNone.Value, acamLeadNONE)                     '..07.18.02 - rg

        .LeadOutType = Switch(optLeadOutLine.Value, acamLeadLINE, _
                              optLeadOutArc.Value, acamLeadARC, _
                              (optLeadOutBoth.Value And chkLeadLineArcTangentialOut.Value), acamLeadBOTH, _
                              (optLeadOutBoth.Value And (Not chkLeadLineArcTangentialOut.Value)), acamLeadBOTH_NOT_TANGENTIAL, _
                              optLeadOutNone.Value, acamLeadNONE)                   '..07.18.02 - rg
        
        .LeadInSloping = chkLeadInSloping.Value
        .LeadOutSloping = chkLeadOutSloping.Value
                                                                  
    End With

Exit Function

'..woops
mbln_SetLeadData_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_SetLeadData"
    mbln_SetLeadData = False

End Function

Private Function mbln_SetLeadDataPocket() As Boolean

On Error GoTo mbln_SetLeadDataPocket_Error

    '..start out fine
    mbln_SetLeadDataPocket = True

    If chkUse3DApproach Then
    
        With clsPathData
                
            '..fill up the lead info
            .LeadLineLengthIn = PDbl(txtZigZagLength.Text)
            .LeadApproachAngle = PDbl(txtApproachAngle.Text)
            .Pocket3DApproach = True
                                                                      
        End With

    Else
    
        clsPathData.Pocket3DApproach = False
    
    End If

Exit Function

'..woops
mbln_SetLeadDataPocket_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_SetLeadDataPocket"
    mbln_SetLeadDataPocket = False

End Function


Private Function mbln_SetMachiningOptionsData_1() As Boolean

On Error GoTo mbln_SetMachiningOptionsData_1_Error
    
    '..start out cool
    mbln_SetMachiningOptionsData_1 = True

    With clsPathData
        
        '..save the depth percentage option to the registry
        SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "IsDepthAPercent", chkDepthPercentage.Value
        
        '..z levels
        .SafeRapidLevel = PDbl(txtSafeRapidLevel.Text)
        .RapidDownTo = PDbl(txtRapidDownTo.Text)
        .MaterialTop = PDbl(txtMaterialTop.Text)
        .FinalDepth = PDbl(txtFinalDepth.Text)
        .FinalDepthPercentage = PDbl(txtFinalDepthPercent.Text)
        .IsFinalDepthPercent = chkDepthPercentage.Value
        .FirstCutDepthPercentage = PDbl(txtThicknessOfFirstCutPercent.Text)
        .LastCutDepthPercentage = PDbl(txtThicknessOfLastCutPercent.Text)
        
        '..multiple passes
        .NumberOfCuts = PDbl(txtNumberOfCuts.Text)
        .DepthOfCutsSpecified = optDepthsOfCutsSpecified.Value
        .ThicknessOfFirstCut = PDbl(txtThicknessOfFirstCut.Text)
        .ThicknessOfLastCut = PDbl(txtThicknessOfLastCut.Text)
        
        If PDbl(txtNumberOfCuts.Text) > 1 Then
            .MultiplePasses = True
        Else
            .MultiplePasses = False
        End If
                    
    End With
    
Exit Function

mbln_SetMachiningOptionsData_1_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_SetMachiningOptionsData_1"
    mbln_SetMachiningOptionsData_1 = False

End Function

Private Function mbln_SetMachiningOptionsData_2() As Boolean

On Error GoTo mbln_SetMachiningOptionsData_2_Error
    
    '..start out cool
    mbln_SetMachiningOptionsData_2 = True

    With clsPathData
        
        '..tooling
        .ToolNumber = PDbl(txtToolNumber.Text)
        .OffsetNumber = PDbl(txtOffsetNumber.Text)
        .SpindleSpeed = PDbl(txtSpindleSpeed.Text)
        .DownFeed = PDbl(txtDownFeed.Text)
        .CutFeed = PDbl(txtCutFeed.Text)
        
        '..machining
        .Stock = PDbl(txtStockToBeLeft.Text)
        .EngraveCornerAngle = PDbl(txtEngravingMaxAngle.Text)
        .WidthOfCut = PDbl(txtWidthOfCut.Text)
        .CutDirection = PDbl(txtCutDirection.Text)
        
        If .MachineMethod = DEF_MACHINE_METHOD_ENGRAVE Then
          .ChordError = PDbl(txtChordError.Text)
          .StepLength = PDbl(txtStepLength.Text)
        ElseIf .MachineMethod = DEF_MACHINE_METHOD_SIMPLE_ENGRAVE Then
          .SimpleEngraveFeed = PDbl(txtChordError.Text)
          .SimpleEngraveClearance = PDbl(txtStepLength.Text)
        End If
        
    End With
    
Exit Function

mbln_SetMachiningOptionsData_2_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_SetMachiningOptionsData_2"
    mbln_SetMachiningOptionsData_2 = False
    
End Function

Private Function mbln_SetMachiningOptionsData_3() As Boolean

On Error GoTo mbln_SetMachiningOptionsData_3_Error
    
    '..start out cool
    mbln_SetMachiningOptionsData_3 = True

    With clsPathData
        .SlowDownForCorners = chkSlowDownForCorners.Value
        .DecelerationDistance = PDbl(txtDecelerationDistance.Text)
        .NumberOfSteps = PDbl(txtNumberOfSteps.Text)
        .SlowDownTo = PDbl(txtSlowDownTo.Text)
        .DoNotSlowDownRadius = PDbl(txtDoNotSlowDownRadius.Text)
        .IgnoreAngleGreaterThan = PDbl(txtIgnoreAngle.Text)
        .AccelerateOutOfCorner = chkAccelerateOutOfCorner.Value
    End With
    
Exit Function

mbln_SetMachiningOptionsData_3_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_SetMachiningOptionsData_3"
    mbln_SetMachiningOptionsData_3 = False
    
End Function


Private Function mbln_CheckMultipleDepths() As Boolean

    Dim dblTotalDepth           As Double
    Dim dblFirstPlusLast        As Double

On Error GoTo mbln_CheckMultipleDepths_Error

    '..check the number of pass
    If clsPathData.NumberOfCuts = 2 Then
        
        '..only 2 passes should be ok by now
        mbln_CheckMultipleDepths = True
        
    Else
        
        '..more than 2 so check
        If chkDepthPercentage.Value Then
            
            '..we are looking at the percentages
            dblFirstPlusLast = PDbl(txtThicknessOfFirstCutPercent.Text) + PDbl(txtThicknessOfLastCutPercent.Text)
                
            If dblFirstPlusLast >= PDbl(txtFinalDepthPercent.Text) Then
                
                '..let the user know what the hell is going on
                MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 70) & Space(3), vbExclamation, DEF_PROJECT_NAME
                 
                '..not good
                mbln_CheckMultipleDepths = False
            
            Else
                
                '..just fine
                mbln_CheckMultipleDepths = True
                
            End If
                
        Else
        
            '..we are looking at the specified depths
            dblTotalDepth = Abs(PDbl(txtMaterialTop.Text) - PDbl(txtFinalDepth.Text))
            dblFirstPlusLast = Abs(PDbl(txtThicknessOfFirstCut.Text) + PDbl(txtThicknessOfLastCut.Text))
                     
            '..make sure the specified thicknesses are not greater than the total depth
            If dblFirstPlusLast >= dblTotalDepth Then
                
                '..let the user know what the hell is going on
                MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 70) & Space(3), vbExclamation, DEF_PROJECT_NAME
                 
                '..not good
                mbln_CheckMultipleDepths = False
            
            Else
                
                '..just fine
                mbln_CheckMultipleDepths = True
                
            End If
            
        End If
                 
    End If
    
Exit Function

mbln_CheckMultipleDepths_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_CheckMultipleDepths"
    mbln_CheckMultipleDepths = False

End Function

Private Function mbln_SetPocketData() As Boolean

On Error GoTo mbln_SetPocketData_Error
    
    '..start out cool
    mbln_SetPocketData = True

    With clsPathData
        
        .MachineMethod = DEF_MACHINE_METHOD_POCKET
        
        '..the pocket type
        .PocketType = Switch(optPocketTypeContour.Value, acamPocketCONTOUR, _
                             optPocketTypeLinear.Value, acamPocketLINEAR, _
                             optPocketTypeSpiral.Value, acamPocketSPIRAL)
                          
        '..look for spiral pocket type
        If .PocketType = acamPocketSPIRAL Then
            MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 89) & vbCrLf & vbCrLf & _
                   App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 90), vbInformation, DEF_PROJECT_NAME
        End If
                        
        '..get final pass around islands
        .FinalPassAroundIslands = Switch(optPocketFinalPassFull.Value, acamFinalPassFULL, _
                                         optPocketFinalPassPartial.Value, acamFinalPassPARTIAL, _
                                         optPocketFinalPassNone.Value, acamFinalPassNONE)
        
        '..get start cutting at
        .StartCutting = Switch(optPocketStartCuttingAtInside.Value, acamStartINSIDE, _
                               optPocketStartCuttingAtOutside.Value, acamStartOUTSIDE)
                               
        '..distance to pocket Boundary
        .PocketBoundary = PDbl(txtPocketBoundary.Text)
        
    End With
    
Exit Function

mbln_SetPocketData_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_SetPocketData"
    mbln_SetPocketData = False

End Function

Private Function mbln_SetRoughFinishData() As Boolean

On Error GoTo mbln_SetRoughFinishData_Error
    
    '..start out cool
    mbln_SetRoughFinishData = True

    With clsPathData
                
        '..compensation
        .MachineCompensation = Switch(optCompAPS.Value, acamCompTOOLCEN, _
                                      optCompMachine.Value, acamCompMC, _
                                      optCompBoth.Value, acamCompBOTH)
        
        '..xy corners
        .XYCorners = Switch(optXYCornersRollRound.Value, acamCornersROUND, _
                            optXYCornersStraight.Value, acamCornersSTRAIGHT)
                                
        '..apply comp on rapid?
        .ApplyCompOnRapid = chkApplyCompOnRapid.Value
        
        '..make sure the is no pocket Boundary distance
        .PocketBoundary = 0
        
        '..full or partial cut?
        .CutType = Switch(optCutTypeFull.Value, DEF_CUT_TYPE_FULL, _
                         optCutTypePartial.Value, DEF_CUT_TYPE_PARTIAL)
        
    End With
    
Exit Function

mbln_SetRoughFinishData_Error:
    
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_SetRoughFinishData"
    mbln_SetRoughFinishData = False

End Function

Private Sub optPocketTypeContour_Click()
    
On Error Resume Next
    
    fraFinalPassAroundIslands.Enabled = True
    fraStartCuttingAt.Enabled = True
    optPocketFinalPassFull.Enabled = True
    optPocketFinalPassNone.Enabled = True
    optPocketFinalPassPartial.Enabled = True
    optPocketStartCuttingAtInside.Enabled = True
    optPocketStartCuttingAtOutside.Enabled = True
    
    If (Err.Number <> 0) Then WriteError Err, False, "optPocketTypeContour_Click"

End Sub

Private Sub optPocketTypeLinear_Click()

On Error Resume Next

    fraFinalPassAroundIslands.Enabled = False
    fraStartCuttingAt.Enabled = False
    optPocketFinalPassFull.Enabled = False
    optPocketFinalPassNone.Enabled = False
    optPocketFinalPassPartial.Enabled = False
    optPocketStartCuttingAtInside.Enabled = False
    optPocketStartCuttingAtOutside.Enabled = False

    If (Err.Number <> 0) Then WriteError Err, False, "optPocketTypeLinear_Click"

End Sub

Private Sub optPocketTypeSpiral_Click()

On Error Resume Next

    fraFinalPassAroundIslands.Enabled = False
    fraStartCuttingAt.Enabled = True
    optPocketFinalPassFull.Enabled = False
    optPocketFinalPassNone.Enabled = False
    optPocketFinalPassPartial.Enabled = False
    optPocketStartCuttingAtInside.Enabled = True
    optPocketStartCuttingAtOutside.Enabled = True
    
    If (Err.Number <> 0) Then WriteError Err, False, "optPocketTypeSpiral_Click"
    
End Sub

Private Sub txtPocketBoundary_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    gbln_TextCalc txtPocketBoundary, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtPocketBoundary_BeforeUpdate"
    
End Sub

Private Sub optCompAPS_Click()

    chkApplyCompOnRapid.Enabled = False

    If (Err.Number <> 0) Then WriteError Err, False, "optCompAPS_Click"
    
End Sub

Private Sub optCompMachine_Click()
    
    chkApplyCompOnRapid.Enabled = True
    
    If (Err.Number <> 0) Then WriteError Err, False, "optCompMachine_Click"
    
End Sub

Private Sub optDepthsOfCutsEqual_Click()
    
On Error Resume Next

    lblThicknessOfFirstCut.Enabled = False
    lblThicknessOfLastCut.Enabled = False
    txtThicknessOfFirstCut.Enabled = False
    txtThicknessOfLastCut.Enabled = False
    txtThicknessOfFirstCutPercent.Enabled = False
    txtThicknessOfLastCutPercent.Enabled = False

    If (Err.Number <> 0) Then WriteError Err, False, "optDepthsOfCutsEqual_Click"
    
End Sub

Private Sub optDepthsOfCutsSpecified_Click()

    Dim bEnabled                As Boolean

On Error Resume Next

    bEnabled = chkDepthPercentage.Value

    lblThicknessOfFirstCut.Enabled = True
    lblThicknessOfLastCut.Enabled = True
    
    txtThicknessOfFirstCut.Enabled = Not bEnabled
    txtThicknessOfLastCut.Enabled = Not bEnabled
    txtThicknessOfFirstCutPercent.Enabled = bEnabled
    txtThicknessOfLastCutPercent.Enabled = bEnabled

    If (Err.Number <> 0) Then WriteError Err, False, "optDepthsOfCutsSpecified_Click"

End Sub

Private Sub txtNumberOfCuts_Change()
        
    Dim bEnabled                As Boolean
            
On Error Resume Next
        
    bEnabled = CBool(1 - PDbl(txtNumberOfCuts.Text))

    optDepthsOfCutsEqual.Enabled = bEnabled
    optDepthsOfCutsSpecified.Enabled = bEnabled
    lblThicknessOfFirstCut.Enabled = bEnabled
    lblThicknessOfLastCut.Enabled = bEnabled
    
    If bEnabled Then
        
        If optDepthsOfCutsSpecified.Value Then
        
            If chkDepthPercentage.Value Then
            
                txtThicknessOfFirstCut.Enabled = Not bEnabled
                txtThicknessOfLastCut.Enabled = Not bEnabled
                txtThicknessOfFirstCutPercent.Enabled = bEnabled
                txtThicknessOfLastCutPercent.Enabled = bEnabled
            
            Else
            
                txtThicknessOfFirstCut.Enabled = bEnabled
                txtThicknessOfLastCut.Enabled = bEnabled
                txtThicknessOfFirstCutPercent.Enabled = Not bEnabled
                txtThicknessOfLastCutPercent.Enabled = Not bEnabled
            
            End If
            
        Else
            
            lblThicknessOfFirstCut.Enabled = False
            lblThicknessOfLastCut.Enabled = False
            txtThicknessOfFirstCut.Enabled = False
            txtThicknessOfLastCut.Enabled = False
            txtThicknessOfFirstCutPercent.Enabled = False
            txtThicknessOfLastCutPercent.Enabled = False
        
        End If
        
    Else
    
        txtThicknessOfFirstCut.Enabled = bEnabled
        txtThicknessOfLastCut.Enabled = bEnabled
        txtThicknessOfFirstCutPercent.Enabled = bEnabled
        txtThicknessOfLastCutPercent.Enabled = bEnabled
    
    End If
    
    DoEvents

    If (Err.Number <> 0) Then WriteError Err, False, "frmWizard_txtNumbersOfCuts_Change"
    
End Sub

Private Sub txtNumberOfCuts_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    '..can't type here!
    KeyCode = 0

End Sub

Private Sub txtNumberOfCuts_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    '..can't type here!
    KeyAscii = 0

End Sub

Private Sub txtSafeRapidLevel_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtSafeRapidLevel, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtSafeRapidLevel_BeforeUpdate"
    
End Sub

Private Sub txtMaterialTop_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtMaterialTop, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtMaterialTop_BeforeUpdate"

End Sub

Private Sub txtFinalDepth_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtFinalDepth, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtFinalDepth_BeforeUpdate"
    
End Sub

Private Sub txtRapidDownTo_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtRapidDownTo, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtRapidDownTo_BeforeUpdate"
    
End Sub

Private Sub txtThicknessOfFirstCut_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim dblDepth            As Double
    
On Error Resume Next
    
    If chkDepthPercentage.Value Then Exit Sub
    
    '..eval value
    If gbln_TextCalc(txtThicknessOfFirstCut, Cancel) Then
    
        If PDbl(txtNumberOfCuts.Text) = 2 Then
        
            dblDepth = PDbl(txtMaterialTop.Text) - PDbl(txtFinalDepth.Text)
            dblDepth = Abs(dblDepth - PDbl(txtThicknessOfFirstCut.Text))
        
            txtThicknessOfLastCut.Text = gs_NoComma(CStr(dblDepth))

        End If

    End If

    If (Err.Number <> 0) Then WriteError Err, False, "txtThicknessOfFirstCut_BeforeUpdate"
    
End Sub

Private Sub txtThicknessOfLastCut_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim dblDepth            As Double
    
On Error Resume Next
    
    If chkDepthPercentage.Value Then Exit Sub
    
    '..eval value
    If gbln_TextCalc(txtThicknessOfLastCut, Cancel) Then
        
        If PDbl(txtNumberOfCuts.Text) = 2 Then
        
            dblDepth = PDbl(txtMaterialTop.Text) - PDbl(txtFinalDepth.Text)
            dblDepth = Abs(dblDepth - PDbl(txtThicknessOfLastCut.Text))
                
            txtThicknessOfFirstCut.Text = gs_NoComma(CStr(dblDepth))
        
        End If
        
    End If
   
    If (Err.Number <> 0) Then WriteError Err, False, "txtThicknessOfLastCut_BeforeUpdate"
   
End Sub

Private Sub txtCutDirection_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtCutDirection, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtCutDirection_BeforeUpdate"

End Sub

Private Sub txtDownFeed_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtDownFeed, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtDownFeed_BeforeUpdate"

End Sub

Private Sub txtCutFeed_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtCutFeed, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtCutFeed_BeforeUpdate"

End Sub

Private Sub txtStockToBeLeft_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtStockToBeLeft, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtStockToBeLeft_BeforeUpdate"

End Sub

Private Sub txtChordError_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtChordError, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtChordError_BeforeUpdate"

End Sub

Private Sub txtStepLength_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtStepLength, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtStepLength_BeforeUpdate"
    
End Sub

Private Sub txtWidthOfCut_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtWidthOfCut, Cancel
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtWidthOfCut_BeforeUpdate"

End Sub

Private Sub txtToolNumber_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    If Not gbln_IsOk(txtToolNumber, Cancel) Then
        
        With txtToolNumber
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    
    End If
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtToolNumber_BeforeUpdate"

End Sub

Private Sub txtSpindleSpeed_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    gbln_TextCalc txtSpindleSpeed, Cancel

    If (Err.Number <> 0) Then WriteError Err, False, "txtSpindleSpeed_BeforeUpdate"

End Sub

Private Sub txtOffsetNumber_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    '..eval value
    If Not gbln_IsOk(txtOffsetNumber, Cancel) Then
    
            With txtOffsetNumber
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    
    End If
    
    If (Err.Number <> 0) Then WriteError Err, False, "txtOffset_BeforeUpdate"

End Sub

Private Function mbln_FillFormCaptions() As Boolean

    Dim sCTX                As String

On Error GoTo mbln_FillFormCaptions_Error

    sCTX = clsOptions.CTXFile

    '..so far so good
    mbln_FillFormCaptions = True

    With App.Frame
        
        Me.Caption = Space(1) & .ReadTextFile(sCTX, 5, 1)
        
        '..buttons
        cmdBack.Caption = .ReadTextFile(sCTX, 10, 1)
        cmdCancel.Caption = .ReadTextFile(sCTX, 10, 2)
        cmdNext.Caption = .ReadTextFile(sCTX, 10, 3)
        cmdBackplot.Caption = .ReadTextFile(sCTX, 10, 4)
        cmdHelp.Caption = .ReadTextFile(sCTX, 10, 11)
              
        '..new door type
        lblDoorTypeComments.Caption = .ReadTextFile(sCTX, 25, 5)
        lblDoorTypeCreatedBy.Caption = .ReadTextFile(sCTX, 25, 4)
        lblDoorTypeName.Caption = .ReadTextFile(sCTX, 25, 3)
        
        '..creation method
        optManualDefine.Caption = .ReadTextFile(sCTX, 27, 2)
        lblManualDefine.Caption = .ReadTextFile(sCTX, 27, 3)
        optMachiningStyle.Caption = .ReadTextFile(sCTX, 27, 4)
        lblMachiningStyleDefine.Caption = .ReadTextFile(sCTX, 27, 5)
        
        '..machining method
        optMachiningMethodRoughFinish.Caption = .ReadTextFile(sCTX, 30, 3)
        lblRoughFinishDescripe.Caption = .ReadTextFile(sCTX, 30, 4) '& space(1) & .ReadTextFile(sCTX, 30, 5)
        optMachiningMethodPocket.Caption = .ReadTextFile(sCTX, 30, 6)
        lblPocketDescribe.Caption = .ReadTextFile(sCTX, 30, 7) '& vbCrLf & .ReadTextFile(sCTX, 30, 8)
        optMachiningMethodEngrave.Caption = .ReadTextFile(sCTX, 30, 9)
        lblEngraveDescribe.Caption = .ReadTextFile(sCTX, 30, 10) '& vbCrLf & .ReadTextFile(sCTX, 30, 11)
        optMachineMethodInsert.Caption = .ReadTextFile(sCTX, 30, 12)
        lblInsertPathDescribe.Caption = .ReadTextFile(sCTX, 30, 13) '& vbCrLf & .ReadTextFile(sCTX, 30, 14)
        optMachiningMethodSimpleEngrave.Caption = .ReadTextFile(sCTX, 30, 14)
        lblSimpleEngraveDescribe.Caption = .ReadTextFile(sCTX, 30, 15)
        
        '..offset amount
        fraDirectionSide.Caption = .ReadTextFile(sCTX, 35, 3)
        optOffsetSideCenter.Caption = .ReadTextFile(sCTX, 35, 4)
        optOffsetSideInside.Caption = .ReadTextFile(sCTX, 35, 5)
        optOffsetSideOutside.Caption = .ReadTextFile(sCTX, 35, 6)
        fraLeadEntryPoint.Caption = .ReadTextFile(sCTX, 35, 7)
        optLeadEntryPointMiddle.Caption = .ReadTextFile(sCTX, 35, 8)
        optLeadEntryPointCorner.Caption = .ReadTextFile(sCTX, 35, 9)
        optLeadEntryPointDrawn.Caption = .ReadTextFile(sCTX, 35, 11)
        lblOffset.Caption = .ReadTextFile(sCTX, 35, 10)
        
        '..insert file parameters
        lblFileToInsert.Caption = .ReadTextFile(sCTX, 40, 3)
        fraInsertRefPoint.Caption = .ReadTextFile(sCTX, 40, 4)
        optRefTopLeft.Caption = .ReadTextFile(sCTX, 40, 5)
        optRefTopRight.Caption = .ReadTextFile(sCTX, 40, 6)
        optRefBottomLeft.Caption = .ReadTextFile(sCTX, 40, 7)
        optRefBottomRight.Caption = .ReadTextFile(sCTX, 40, 8)
        optRefLeft.Caption = .ReadTextFile(sCTX, 40, 9)
        optRefRight.Caption = .ReadTextFile(sCTX, 40, 10)
        optRefTop.Caption = .ReadTextFile(sCTX, 40, 11)
        optRefBottom.Caption = .ReadTextFile(sCTX, 40, 12)
        optRefParametric.Caption = .ReadTextFile(sCTX, 40, 18)
        chkInsertFileCenterX.Caption = .ReadTextFile(sCTX, 40, 13)
        chkInsertFileCenterY.Caption = .ReadTextFile(sCTX, 40, 14)
        fraInsertPositionPoint.Caption = .ReadTextFile(sCTX, 40, 15)
        lblInsertPointX.Caption = .ReadTextFile(sCTX, 40, 16)
        lblInsertPointY.Caption = .ReadTextFile(sCTX, 40, 17)
        
        '..tool side and direction
        fraToolSideClosed.Caption = .ReadTextFile(sCTX, 45, 7)
        optDirectionCW.Caption = .ReadTextFile(sCTX, 45, 4)
        optDirectionCCW.Caption = .ReadTextFile(sCTX, 45, 5)
        optDirectionReversed.Caption = .ReadTextFile(sCTX, 45, 6)
        fraToolDirection.Caption = .ReadTextFile(sCTX, 45, 3)
        optSideOutside.Caption = .ReadTextFile(sCTX, 45, 8)
        optSideInside.Caption = .ReadTextFile(sCTX, 45, 9)
        optSideCentreClosed.Caption = .ReadTextFile(sCTX, 45, 10)
        
        fraToolsideOpen.Caption = .ReadTextFile(sCTX, 45, 17)
        optSideLeft.Caption = .ReadTextFile(sCTX, 45, 11)
        optSideRight.Caption = .ReadTextFile(sCTX, 45, 12)
        optSideCentreOpen.Caption = .ReadTextFile(sCTX, 45, 10)

        chkReverseToolDirection.Caption = .ReadTextFile(sCTX, 45, 13)
        
        chkViewGhostTools.Caption = .ReadTextFile(sCTX, 45, 14)
                
        '..pocketing parameters
        fraPocketType.Caption = .ReadTextFile(sCTX, 50, 2)
        optPocketTypeContour.Caption = .ReadTextFile(sCTX, 50, 3)
        optPocketTypeLinear.Caption = .ReadTextFile(sCTX, 50, 4)
        optPocketTypeSpiral.Caption = .ReadTextFile(sCTX, 50, 5)
        fraFinalPassAroundIslands.Caption = .ReadTextFile(sCTX, 50, 6)
        optPocketFinalPassFull.Caption = .ReadTextFile(sCTX, 50, 7)
        optPocketFinalPassPartial.Caption = .ReadTextFile(sCTX, 50, 8)
        optPocketFinalPassNone.Caption = .ReadTextFile(sCTX, 50, 9)
        fraStartCuttingAt.Caption = .ReadTextFile(sCTX, 50, 10)
        optPocketStartCuttingAtInside.Caption = .ReadTextFile(sCTX, 50, 11)
        optPocketStartCuttingAtOutside.Caption = .ReadTextFile(sCTX, 50, 12)
        lblPocketBoundary.Caption = .ReadTextFile(sCTX, 50, 13)
        
        '..rough/finish parameters
        fraCompensationType.Caption = .ReadTextFile(sCTX, 55, 2)
        optCompAPS.Caption = .ReadTextFile(sCTX, 55, 3)
        optCompBoth.Caption = .ReadTextFile(sCTX, 55, 4)
        optCompMachine.Caption = .ReadTextFile(sCTX, 55, 5)
        chkApplyCompOnRapid.Caption = .ReadTextFile(sCTX, 55, 6)
        fraXYCorners.Caption = .ReadTextFile(sCTX, 55, 7)
        optXYCornersRollRound.Caption = .ReadTextFile(sCTX, 55, 8)
        optXYCornersStraight.Caption = .ReadTextFile(sCTX, 55, 9)
        fraCutType.Caption = .ReadTextFile(sCTX, 55, 10)
        optCutTypeFull.Caption = .ReadTextFile(sCTX, 55, 11)
        optCutTypePartial.Caption = .ReadTextFile(sCTX, 55, 12)
        
        '..machining options 1
        fraZLevels.Caption = .ReadTextFile(sCTX, 60, 2)
        lblSafeRapidLevel.Caption = .ReadTextFile(sCTX, 60, 3)
        lblMaterialTop.Caption = .ReadTextFile(sCTX, 60, 4)
        lblRapidDownTo.Caption = .ReadTextFile(sCTX, 60, 5)
        lblFinalDepth.Caption = .ReadTextFile(sCTX, 60, 6)
        fraDepthsOfCuts.Caption = .ReadTextFile(sCTX, 60, 7)
        optDepthsOfCutsEqual.Caption = .ReadTextFile(sCTX, 60, 8)
        optDepthsOfCutsSpecified.Caption = .ReadTextFile(sCTX, 60, 9)
        lblNumberOfCuts.Caption = .ReadTextFile(sCTX, 60, 10)
        lblThicknessOfFirstCut.Caption = .ReadTextFile(sCTX, 60, 11)
        lblThicknessOfLastCut.Caption = .ReadTextFile(sCTX, 60, 12)
        chkDepthPercentage.Caption = .ReadTextFile(sCTX, 60, 13)
        
        '..machining options 2
        fraTooling.Caption = .ReadTextFile(sCTX, 65, 2)
        lblToolNumber.Caption = .ReadTextFile(sCTX, 65, 3)
        lblDiameter.Caption = .ReadTextFile(sCTX, 65, 4)
        lblDownFeed.Caption = .ReadTextFile(sCTX, 65, 5)
        lblOffsetNumber.Caption = .ReadTextFile(sCTX, 65, 6)
        lblSpindleSpeed.Caption = .ReadTextFile(sCTX, 65, 7)
        lblCutFeed.Caption = .ReadTextFile(sCTX, 65, 8)
        fraMachining.Caption = .ReadTextFile(sCTX, 65, 9)
        lblStockToLeave.Caption = .ReadTextFile(sCTX, 65, 10)
        'lblChordError.Caption = .ReadTextFile(sCTX, 65, 11)
        lblEngravingMaxAngle.Caption = .ReadTextFile(sCTX, 65, 15)
        'lblStepLength.Caption = .ReadTextFile(sCTX, 65, 12)
        lblWidthOfCut.Caption = .ReadTextFile(sCTX, 65, 13)
        lblCutDirection.Caption = .ReadTextFile(sCTX, 65, 14)


        '..machining options 3
        chkSlowDownForCorners.Caption = .ReadTextFile(sCTX, 67, 2)
        fraSlowDownForCorners.Caption = .ReadTextFile(sCTX, 67, 3)
        lblDecelerationDistance.Caption = .ReadTextFile(sCTX, 67, 4)
        lblNumberOfSteps.Caption = .ReadTextFile(sCTX, 67, 5)
        lblSlowDownTo.Caption = .ReadTextFile(sCTX, 67, 6)
        lblDoNotSlowDownRadius.Caption = .ReadTextFile(sCTX, 67, 7)
        lblIgnoreAngle.Caption = .ReadTextFile(sCTX, 67, 8)
        chkAccelerateOutOfCorner.Caption = .ReadTextFile(sCTX, 67, 9)


        '..lead info
        fraLeadIn.Caption = .ReadTextFile(sCTX, 70, 2)
        fraLeadOut.Caption = .ReadTextFile(sCTX, 70, 3)
        optLeadInLine.Caption = .ReadTextFile(sCTX, 70, 4)
        optLeadOutLine.Caption = .ReadTextFile(sCTX, 70, 4)
        optLeadInBoth.Caption = .ReadTextFile(sCTX, 70, 5)
        optLeadOutBoth.Caption = .ReadTextFile(sCTX, 70, 5)
        optLeadInArc.Caption = .ReadTextFile(sCTX, 70, 6)
        optLeadOutArc.Caption = .ReadTextFile(sCTX, 70, 6)
        optLeadInNone.Caption = .ReadTextFile(sCTX, 70, 7)
        optLeadOutNone.Caption = .ReadTextFile(sCTX, 70, 7)
        chkLeadInSloping.Caption = .ReadTextFile(sCTX, 70, 8)
        chkLeadOutSloping.Caption = .ReadTextFile(sCTX, 70, 8)
        lblLeadLineLengthIn.Caption = .ReadTextFile(sCTX, 70, 9)                    '..07.18.02 - rg
        lblLeadLineLengthOut.Caption = .ReadTextFile(sCTX, 70, 9)                   '..07.18.02 - rg
        chkLeadLineArcTangentialIn.Caption = .ReadTextFile(sCTX, 70, 14)            '..07.18.02 - rg
        chkLeadLineArcTangentialOut.Caption = .ReadTextFile(sCTX, 70, 14)           '..07.18.02 - rg
        lblLeadArcRadius.Caption = .ReadTextFile(sCTX, 70, 10)
        lblLeadApproachAngle.Caption = .ReadTextFile(sCTX, 70, 11)
        lblLeadOverlap.Caption = .ReadTextFile(sCTX, 70, 12)
        lblTip.Caption = .ReadTextFile(sCTX, 70, 13)
        
        '..3D Approach
        chkUse3DApproach.Caption = .ReadTextFile(sCTX, 72, 2)
        fra3DApproachParameters.Caption = .ReadTextFile(sCTX, 72, 3)
        lblApproachAngle.Caption = .ReadTextFile(sCTX, 72, 4)
        lblZigZagLength.Caption = .ReadTextFile(sCTX, 72, 5)
        
        '..Machining styles
        fraMachiningStyles.Caption = .ReadTextFile(sCTX, 73, 1)
        lblSelectMachiningStyle.Caption = .ReadTextFile(sCTX, 73, 2)
        optCutTypeFull.Caption = .ReadTextFile(sCTX, 55, 11)
        optCutTypePartial.Caption = .ReadTextFile(sCTX, 55, 12)
        optStylesCutTypeFull.Caption = .ReadTextFile(sCTX, 55, 11)
        optStylesCutTypePartial.Caption = .ReadTextFile(sCTX, 55, 12)
        
        '..path complete
        lblPathComplete.Caption = .ReadTextFile(sCTX, 75, 2)
        
    End With

Controlled_Exit:

Exit Function

mbln_FillFormCaptions_Error:

    If (Err.Number <> 0) Then WriteError Err, True, "mbln_FillFormCaptions"
    mbln_FillFormCaptions = False
    Resume Controlled_Exit

End Function

Private Function mbln_CheckAllText(Ctl As Control) As Boolean

    Dim i                   As Integer
    
On Error Resume Next

    'start out cool
    mbln_CheckAllText = True
    
    For i = 0 To MultiPage1.Pages.Count
        
        If MultiPage1.Value = i Then
        
            'check for empty text boxes and warn user if any
            For Each Ctl In MultiPage1.Pages(i).Controls
                
                '..check for textbox only
                If TypeOf Ctl Is TextBox Then
                    
                    '..check for enabled
                    If (Ctl.Enabled And Ctl.Visible) Then
                    
                        'check for empty
                        If Len(Ctl.Text) = 0 Then
                            
                            'woops ~ something missing :(
                            MsgBox App.Frame.ReadTextFile(clsOptions.CTXFile, 500, 80) & Space(3), vbExclamation, DEF_PROJECT_NAME
                            mbln_CheckAllText = False
                            Ctl.SetFocus
                            DoEvents
                            Exit Function
                            
                        End If
                    
                    End If
                            
                End If

            Next Ctl
            
        End If
    
    Next i

    If (Err.Number <> 0) Then WriteError Err, False, "mbln_CheckAllText"

End Function

Private Function mint_CheckNullToolSide(ByRef vValue As Variant) As Integer
    
    If IsNull(vValue) Then
        mint_CheckNullToolSide = 0
    Else
        mint_CheckNullToolSide = CInt(vValue)
    End If

End Function


