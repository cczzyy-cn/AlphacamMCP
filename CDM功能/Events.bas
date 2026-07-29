Option Explicit

Public m_AlphaDOORRunning As Boolean

Dim sCurrentDB As String


Private Function mbln_UpdateOperatorDescriptions() As Boolean
  Dim strOperator1  As String
  Dim strOperator2  As String
  Dim strOperator3  As String
  Dim strOperator4  As String
  Dim strOperator5  As String
  Dim strOperator6  As String
  Dim strCTX        As String
  Dim strSQL        As String
  Dim lngRet        As Long
  Dim lngTot        As Long
  Dim clsOptions    As New COptions
'
  strCTX = clsOptions.CTXFile
  
  With Frame
    strOperator1 = .ReadTextFile(strCTX, 186, 1)
    strOperator2 = .ReadTextFile(strCTX, 186, 2)
    strOperator3 = .ReadTextFile(strCTX, 186, 3)
    strOperator4 = .ReadTextFile(strCTX, 186, 4)
    strOperator5 = .ReadTextFile(strCTX, 186, 5)
    strOperator6 = .ReadTextFile(strCTX, 186, 6)
  End With
  
  strSQL = "UPDATE AD_OPERATORS SET OperatorDescription='" & gs_FixSQL(strOperator1) & "' WHERE OperatorID=1"
  gdb_CDM.Execute strSQL, lngRet
  lngTot = lngTot + lngRet
  
  strSQL = "UPDATE AD_OPERATORS SET OperatorDescription='" & gs_FixSQL(strOperator2) & "' WHERE OperatorID=2"
  gdb_CDM.Execute strSQL, lngRet
  lngTot = lngTot + lngRet
  
  strSQL = "UPDATE AD_OPERATORS SET OperatorDescription='" & gs_FixSQL(strOperator3) & "' WHERE OperatorID=3"
  gdb_CDM.Execute strSQL, lngRet
  lngTot = lngTot + lngRet
  
  strSQL = "UPDATE AD_OPERATORS SET OperatorDescription='" & gs_FixSQL(strOperator4) & "' WHERE OperatorID=4"
  gdb_CDM.Execute strSQL, lngRet
  lngTot = lngTot + lngRet
  
  strSQL = "UPDATE AD_OPERATORS SET OperatorDescription='" & gs_FixSQL(strOperator5) & "' WHERE OperatorID=5"
  gdb_CDM.Execute strSQL, lngRet
  lngTot = lngTot + lngRet
  
  strSQL = "UPDATE AD_OPERATORS SET OperatorDescription='" & gs_FixSQL(strOperator6) & "' WHERE OperatorID=6"
  gdb_CDM.Execute strSQL, lngRet
  lngTot = lngTot + lngRet
  
  ' Success if all operators have been modified
  mbln_UpdateOperatorDescriptions = lngTot = 6
  
End Function

Private Function mbln_InsertOperators() As Boolean
'
  On Error GoTo mbln_InsertOperators_Error
    
  ' Assume success
  mbln_InsertOperators = True
  
  ' Insert Operators
  gdb_CDM.Execute ("INSERT INTO AD_OPERATORS(Operator, OperatorDescription) VALUES ('=','Equal To')")
  gdb_CDM.Execute ("INSERT INTO AD_OPERATORS(Operator, OperatorDescription) VALUES ('>','Greater Than')")
  gdb_CDM.Execute ("INSERT INTO AD_OPERATORS(Operator, OperatorDescription) VALUES ('<','Less Than')")
  gdb_CDM.Execute ("INSERT INTO AD_OPERATORS(Operator, OperatorDescription) VALUES ('>=','Greater Than or Equal To')")
  gdb_CDM.Execute ("INSERT INTO AD_OPERATORS(Operator, OperatorDescription) VALUES ('<=','Less Than or Equal To')")
  gdb_CDM.Execute ("INSERT INTO AD_OPERATORS(Operator, OperatorDescription) VALUES ('<>','Not Equal To')")
  
Exit Function
  
mbln_InsertOperators_Error:

  mbln_InsertOperators = False
  MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME

End Function



Private Function mbln_TestUpdateOperators() As Boolean
  Dim rstOperators    As ADODB.Recordset
  Dim strCTXOperator1 As String
  Dim strDBOperator1  As String
'
  Set rstOperators = grst_GetAllOperators
  
  If Not rstOperators Is Nothing Then
    strCTXOperator1 = Frame.ReadTextFile(strCTX, 186, 1)
    
    rstOperators.MoveFirst
    strDBOperator1 = rstOperators.Fields!OperatorDescription
    
    If UCase(strCTXOperator1) <> UCase(strDBOperator1) Then
      mbln_TestUpdateOperators = True
    End If
    
    rstOperators.Close
    
  End If

  Set rstOperators = Nothing

End Function


Public Sub m_OrderToolPaths()
  Dim sCTX As String
'
  Set clsOptions = New COptions
  sCTX = clsOptions.CTXFile
  
  With frmOrderNestedToolpaths
    .bCancel = False
    .txtLastTool = GetSetting("Licom Systems", "OrderNestedToolpaths", "LastTool", "")
    .txtLastTool.Tag = GetSetting("Licom Systems", "OrderNestedToolpaths", "LastToolFilename", "")
    If .txtLastTool.Tag = "" Then
      .txtLastTool = ""
    End If
    .Show
    If Not .bCancel Then
      OrderToolpaths
    End If
  End With
End Sub

Public Function mbln_BackupDatabase(NumberOfBackups As Integer) As Boolean
  Dim sini              As String
  Dim dbINI             As New CIniFile
  Dim sDatabasePath     As String
  Dim sDatabaseFile     As String
  Dim iNextBackupNum    As Integer
  Dim sCTX              As String
  
  On Error GoTo EH
  
  With Frame
    sCTX = .PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT
    .ShowProgressBox .ReadTextFile(sCTX, 110, 1), .ReadTextFile(sCTX, 300, 25)
  End With
  
  sDatabasePath = gstr_ParseDir(sCurrentDB)
  sDatabaseFile = gstr_ParseName(sCurrentDB)
  
  sini = LicomdatPath & "Licomdat\CDM Data\CDM DB Backup.ini"
  
  ' Initialise the class
  dbINI.Init sini
  
  ' Get the next backup number from the INI file
  iNextBackupNum = dbINI.GetNumericValue("BackupCounter", "NextBackupNumber")
  
  FileCopy sCurrentDB, LicomdatPath & "Licomdat\CDM Data\CDM_Backup" & iNextBackupNum & ".mdb"
  
  iNextBackupNum = iNextBackupNum + 1
  
  If iNextBackupNum > NumberOfBackups Then
    iNextBackupNum = 1
  End If
  
  dbINI.SaveNumericValue "BackupCounter", "NextBackupNumber", iNextBackupNum

  mbln_BackupDatabase = True

EH:
  Frame.CloseProgressBox
End Function

Function InitAlphacamAddIn(AcamVersion As Long) As Integer
        
    Dim sCTX            As String
    Dim lButton         As Long
    Dim FSO             As New Scripting.FileSystemObject
    Dim strMenuName     As String
    Dim blnRet          As Boolean
    Dim IsRibbon        As Boolean

On Error Resume Next
    
    ' 16 mar 11 TFS#43447
    '   + UPDATED to use AddMenuItem32 to support customizable toolbar
    '   + REMOVED mbln_ValidReference as it will never be called, error is raised before it gets that far
    
    With App.Frame
        
        sCTX = .PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT
        
        strMenuName = .ReadTextFile(sCTX, 3, 1)
        
        '..save the alaphacam level to the registry
        SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "AlphaCAM Level", App.ProgramLevel
                
        '..save the path to this macro in the registry
        SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "MacroPath", .PathOfThisAddin
                
        '..save the path to this macro in the registry
        SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "MacroVersion", DEF_MACRO_VERSION

        lButton = .CreateButtonBar(strMenuName)
        
        IsRibbon = .RibbonBarInterface
        
        '..set the menu items
        If .AddMenuItem32(.ReadTextFile(sCTX, 3, 3), "m_Processing", acamMenuNEW, strMenuName, vbNullString, 1) Then
                If FSO.FileExists(.PathOfThisAddin & DEF_BACKSLASH & "main.bmp") Then
                        Call .AddButton(lButton, "main.bmp", .LastMenuCommandID)
                End If
        End If
        
        blnRet = .AddMenuItem32("-", "", acamMenuNEW, strMenuName, vbNullString, 2)
        
        blnRet = .AddMenuItem32(.ReadTextFile(sCTX, 3, 12), "m_TestUserStyles", acamMenuNEW, strMenuName, vbNullString, 3)
        If (IsRibbon And blnRet) Then
            If FSO.FileExists(.PathOfThisAddin & DEF_BACKSLASH & "TestUserStyles.bmp") Then
                        Call .AddButton(lButton, "TestUserStyles.bmp", .LastMenuCommandID)
                End If
        End If
        
        blnRet = .AddMenuItem32(.ReadTextFile(sCTX, 3, 13), "m_OrderToolPaths", acamMenuNEW, strMenuName, vbNullString, 4)
        If (IsRibbon And blnRet) Then
            If FSO.FileExists(.PathOfThisAddin & DEF_BACKSLASH & "OrderToolPaths.bmp") Then
                        Call .AddButton(lButton, "OrderToolPaths.bmp", .LastMenuCommandID)
                End If
        End If
        
        blnRet = .AddMenuItem32("-", "", acamMenuNEW, strMenuName, vbNullString, 5)
        
        ' 15 dec 08 - rg
        '
        If .AddMenuItem32(.ReadTextFile(sCTX, 3, 14), "g_ConfigDB", acamMenuNEW, strMenuName, vbNullString, 6) Then
                
                If FSO.FileExists(.PathOfThisAddin & DEF_BACKSLASH & "udl.bmp") Then
                        Call .AddButton(lButton, "udl.bmp", .LastMenuCommandID)
                End If
                
                blnRet = True
                
        End If
        
        ' 16 mar 11 TFS#43440
        '
        ' add compact db menu item
        If .AddMenuItem32(.ReadTextFile(sCTX, 3, 15), "g_CompactDB", acamMenuNEW, strMenuName, vbNullString, 7) Then
            If (IsRibbon) Then
                If FSO.FileExists(.PathOfThisAddin & DEF_BACKSLASH & "CompactDB.bmp") Then
                    Call .AddButton(lButton, "CompactDB.bmp", .LastMenuCommandID)
                End If
            End If
                ' can probably do without a button for this command
                
                'If FSO.FileExists(gs_ThisDir & "db.bmp") Then
                '        Call .AddButton(lButton, "db.bmp", .LastMenuCommandID)
                'End If
                
                blnRet = True
                
        End If
        '
        ' add support util menu item
        If .AddMenuItem32(.ReadTextFile(sCTX, 3, 16), "g_SupportUtil", acamMenuNEW, strMenuName, vbNullString, 8) Then
                
                If FSO.FileExists(gs_ThisDir & "support.bmp") Then
                        Call .AddButton(lButton, "support.bmp", .LastMenuCommandID)
                End If
                
                blnRet = True
                
        End If
        
        ' Auto Import & Nest
        blnRet = .AddMenuItem32("自动化生产排版", "m_AutoImportNest", acamMenuNEW, strMenuName, vbNullString, 20)
        
        If blnRet Then blnRet = .AddMenuItem32("-", "", acamMenuNEW, strMenuName, vbNullString, 9)
        
        blnRet = .AddMenuItem32(.ReadTextFile(sCTX, 3, 8), "m_Help", acamMenuNEW, strMenuName, vbNullString, 10)
        If (IsRibbon And blnRet) Then
            If FSO.FileExists(.PathOfThisAddin & DEF_BACKSLASH & "Help.bmp") Then
                        Call .AddButton(lButton, "Help.bmp", .LastMenuCommandID)
            End If
        End If
        ' 16 mar 11 TFS#43438
        '   + REMOVED to help declutter the toolbar
        '
        'If FSO.FileExists(.PathOfThisAddin & DEF_BACKSLASH & "help.bmp") Then .AddButton lButton, "help.bmp", .LastMenuCommandID
                
        blnRet = .AddMenuItem32(.ReadTextFile(sCTX, 3, 7), "m_About", acamMenuNEW, strMenuName, vbNullString, 11)
        If (IsRibbon And blnRet) Then
            If FSO.FileExists(.PathOfThisAddin & DEF_BACKSLASH & "About.bmp") Then
                        Call .AddButton(lButton, "About.bmp", .LastMenuCommandID)
            End If
        End If
                            
    End With
    
Controlled_Exit:
    
    Set FSO = Nothing
    
    InitAlphacamAddIn = 0
            
    If (Err.Number <> 0) Then WriteError Err, False, "InitAlphacamAddIn": Err.Clear
            
End Function

' 16 mar 11 TFS#43440
'
Public Sub g_CompactDB()
        
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        ' 05 sep 11 TFS#45900
        '   + CHANGED .exe file name
        '
        blnRet = App.ShellAndWait(App.Frame.PathOfThisAddin & "\CDM_Compress.exe")
        
Controlled_Exit:

Exit Sub

ErrTrap:

        MsgBox Err.Description, vbExclamation
        Resume Controlled_Exit
        
End Sub
'
Public Function OnUpdateg_CompactDB()
        
On Error Resume Next

        ' 05 sep 11 TFS#45900
        '   + CHANGED .exe file name
        '
        OnUpdateg_CompactDB = Abs(CBool(Len(Dir$(App.Frame.PathOfThisAddin & "\CDM_Compress.exe"))))
        
End Function

' 16 mar 11 TFS#43440
'
Public Sub g_SupportUtil()
        
        Dim blnRet                  As Boolean
        
On Error GoTo ErrTrap
        
        blnRet = App.ShellAndWait(App.Frame.PathOfThisAddin & "\CDM_Support.exe")
        
Controlled_Exit:

Exit Sub

ErrTrap:

        MsgBox Err.Description, vbExclamation
        Resume Controlled_Exit
        
End Sub
'
Public Function OnUpdateg_SupportUtil()
        
On Error Resume Next

        OnUpdateg_SupportUtil = Abs(CBool(Len(Dir$(App.Frame.PathOfThisAddin & "\CDM_Support.exe"))))
        
End Function

Public Sub m_TestUserStyles()
  Dim sCTX As String
'
  Set clsOptions = New COptions
  sCTX = clsOptions.CTXFile
  
  With frmUserStyleTest
    DoEvents
    Frame.ShowProgressBox Frame.ReadTextFile(sCTX, 400, 1), Frame.ReadTextFile(sCTX, 400, 16) & " ..."
    LoadDoorStyles
    Frame.CloseProgressBox
    DoEvents
    .Show vbModeless
  End With
  
End Sub

' REMOVED AND REPLACED WITH mint_UpdateDB   07/25/03 - rg
'Private Function m_UpdateEngraving() As Integer
'  Dim rsDoorPaths As ADODB.Recordset
'  Dim lFieldCount As Long
''
'  On Error GoTo eh
'  m_UpdateEngraving = 1
'  Set rsDoorPaths = grst_GetAllPaths
'  If rsDoorPaths Is Nothing Then Exit Function
'
'  lFieldCount = rsDoorPaths.Fields.Count
'  rsDoorPaths.Close
'
'  If lFieldCount = 60 Then
'    ' Update AD_DOOR_PATHS to add Engrave Corner Angle Field
'    gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD EngraveCornerAngle DOUBLE"
'    m_UpdateEngraving = 2
'    ' Change the data type for LeadLineLengthOut from LongInt to Double
'    gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ALTER COLUMN LeadLineLengthOut DOUBLE"
'  End If
'
'  ' Set a return value of 0 to indicate success
'  m_UpdateEngraving = 0
'Exit Function
'
'eh:
'End Function

Private Function mint_UpdateDB() As AdoorUpdateDatabaseFlag
  
    Dim r                       As ADODB.Recordset

On Error GoTo EH
    
    ' !! DON'T FORGET TO UPDATE DEF_DB_VERSION WHEN ADDING NEW DB CONTENT !!

    ' set init return flag
    mint_UpdateDB = adoorUPDATE_DB_ENGRAVE_CORNER
    
    ' get all paths from the database
    Set r = grst_GetAllPaths
    
    If Not (r Is Nothing) Then
        
        ' add EngraveCornerAngle field if not there
        If Not mbln_DBFieldExists(r, "EngraveCornerAngle") Then
        
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD EngraveCornerAngle FLOAT"
            Else
                ' Update AD_DOOR_PATHS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD EngraveCornerAngle DOUBLE"
            End If
            
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_LEADOUT
            
            If gbln_SQLServer Then
                
                g_DropConstraint "ALTER TABLE [dbo].[AD_DOOR_PATHS] DROP CONSTRAINT [DF__AD_DOOR_P__LeadL__32E0915F]"
            
                ' Change the data type for LeadLineLengthOut from LongInt to Double
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ALTER COLUMN LeadLineLengthOut FLOAT"
            Else
                ' Change the data type for LeadLineLengthOut from LongInt to Double
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ALTER COLUMN LeadLineLengthOut DOUBLE"
            End If
            
        End If
        
        ' add Pocket3DApproach field if not there
        If Not mbln_DBFieldExists(r, "Pocket3DApproach") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_POCKET_3DAPPROACH
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the Pocketing 3D Approach lead-in Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD Pocket3DApproach BIT"
            Else
                ' Update AD_DOOR_PATHS to add the Pocketing 3D Approach lead-in Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD Pocket3DApproach YESNO"
            End If
                        
        End If
        
        ' add CreationMethod field if not there
        If Not mbln_DBFieldExists(r, "CreationMethod") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_CREATION_METHOD
            
            ' Update AD_DOOR_PATHS to add the CreationMethod Field
            gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD CreationMethod VARCHAR(50)"
                        
        End If
        
        ' add MachiningStyle field if not there
        If Not mbln_DBFieldExists(r, "MachiningStyle") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_MACHINING_STYLE
            
            ' Update AD_DOOR_PATHS to add the MachiningStyle Field
            gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD MachiningStyle VARCHAR(255)"
                        
        End If
        
        ' add CutType field if not there
        If Not mbln_DBFieldExists(r, "CutType") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_CUT_TYPE
            
            ' Update AD_DOOR_PATHS to add the CutType Field
            gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD CutType VARCHAR(50)"
                        
        End If
    
        ' add PartialStartElemIndex field if not there
        If Not mbln_DBFieldExists(r, "PartialStartElemIndex") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_PARTIAL_START_ELEM_INDEX
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the PartialStartElemIndex Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD PartialStartElemIndex INT"
            Else
                ' Update AD_DOOR_PATHS to add the PartialStartElemIndex Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD PartialStartElemIndex LONG"
            End If
                        
        End If
    
        ' add PartialStartElemDist field if not there
        If Not mbln_DBFieldExists(r, "PartialStartElemDist") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_PARTIAL_START_ELEM_DIST
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the PartialStartElemDist Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD PartialStartElemDist FLOAT"
            Else
                ' Update AD_DOOR_PATHS to add the PartialStartElemDist Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD PartialStartElemDist DOUBLE"
            End If
                        
        End If
    
        ' add PartialEndElemIndex field if not there
        If Not mbln_DBFieldExists(r, "PartialEndElemIndex") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_PARTIAL_END_ELEM_INDEX
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the PartialEndElemIndex Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD PartialEndElemIndex INT"
            Else
                ' Update AD_DOOR_PATHS to add the PartialEndElemIndex Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD PartialEndElemIndex LONG"
            End If
                        
        End If
    
        ' add PartialEndElemDist field if not there
        If Not mbln_DBFieldExists(r, "PartialEndElemDist") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_PARTIAL_END_ELEM_DIST
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the PartialEndElemDist Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD PartialEndElemDist FLOAT"
            Else
                ' Update AD_DOOR_PATHS to add the PartialEndElemDist Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD PartialEndElemDist DOUBLE"
            End If
                        
        End If
        
        ' add SlowDownForCorners field if not there
        If Not mbln_DBFieldExists(r, "SlowDownForCorners") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_SLOW_DOWN_FOR_CORNERS
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the SlowDownForCorners Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD SlowDownForCorners BIT"
            Else
                ' Update AD_DOOR_PATHS to add the SlowDownForCorners Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD SlowDownForCorners YESNO"
            End If
            
        End If
    
        ' add DecelerationDistance field if not there
        If Not mbln_DBFieldExists(r, "DecelerationDistance") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_SLOW_DOWN_DECEL_DIST
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the DecelerationDistance Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD DecelerationDistance FLOAT"
            Else
                ' Update AD_DOOR_PATHS to add the DecelerationDistance Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD DecelerationDistance DOUBLE"
            End If
            
        End If
    
        ' add NumberOfSteps field if not there
        If Not mbln_DBFieldExists(r, "NumberOfSteps") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_SLOW_DOWN_NUM_STEPS
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the NumberOfSteps Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD NumberOfSteps INT"
            Else
                ' Update AD_DOOR_PATHS to add the NumberOfSteps Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD NumberOfSteps LONG"
            End If
            
        End If
    
        ' add SlowDownTo field if not there
        If Not mbln_DBFieldExists(r, "SlowDownTo") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_SLOW_DOWN_TO
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the SlowDownTo Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD SlowDownTo FLOAT"
            Else
                ' Update AD_DOOR_PATHS to add the SlowDownTo Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD SlowDownTo DOUBLE"
            End If
            
        End If
    
        ' add DoNotSlowDownRadius field if not there
        If Not mbln_DBFieldExists(r, "DoNotSlowDownRadius") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_SLOW_DOWN_RADIUS
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the DoNotSlowDownRadius Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD DoNotSlowDownRadius FLOAT"
            Else
                ' Update AD_DOOR_PATHS to add the DoNotSlowDownRadius Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD DoNotSlowDownRadius DOUBLE"
            End If
            
        End If
    
        ' add IgnoreAngleGreaterThan field if not there
        If Not mbln_DBFieldExists(r, "IgnoreAngleGreaterThan") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_SLOW_DOWN_ANGLE
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the IgnoreAngleGreaterThan Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD IgnoreAngleGreaterThan FLOAT"
            Else
                ' Update AD_DOOR_PATHS to add the IgnoreAngleGreaterThan Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD IgnoreAngleGreaterThan DOUBLE"
            End If
        End If
    
        ' add AccelerateOutOfCorner field if not there
        If Not mbln_DBFieldExists(r, "AccelerateOutOfCorner") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_SLOW_DOWN_ACCEL_OUT
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the AccelerateOutOfCorner Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD AccelerateOutOfCorner BIT"
            Else
                ' Update AD_DOOR_PATHS to add the AccelerateOutOfCorner Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD AccelerateOutOfCorner YESNO"
            End If
                        
        End If
    
        ' add InsertParametricGroupNumber field if not there
        If Not mbln_DBFieldExists(r, "InsertParametricGroupNumber") Then
        
            ' set init return flag
            mint_UpdateDB = adoorUPDATE_DB_PATHS_INSERT_PARAMETRIC_GROUP_NUMBER
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add the InsertParametricGroupNumber Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD InsertParametricGroupNumber INT"
            Else
                ' Update AD_DOOR_PATHS to add the InsertParametricGroupNumber Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD InsertParametricGroupNumber LONG"
            End If
        
        End If

        ' add ToolSidePartialReverse field if not there
        If Not mbln_DBFieldExists(r, "ToolSidePartialReverse") Then
            
            mint_UpdateDB = adoorUPDATE_DB_DOOR_PATHS_TOOLSIDE_PARTIAL_REVERSE

            ' Update AWD_DOOR_PATHS to add ToolSidePartialReverse Field
            gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD ToolSidePartialReverse BIT"
        End If
        
        mint_UpdateDB = adoorUPDATE_DB_DOOR_PATHS_START_POINT_AS_DRAWN

        ' Change the type of LeadEntryPointIsCorner from boolean to integer
        ' to accomodate the extra value (leave start point as defined)
        If gbln_SQLServer Then
            g_DropConstraint "ALTER TABLE [dbo].[AD_DOOR_PATHS] DROP CONSTRAINT [DF__AD_DOOR_P__LeadE__36B12243]"
        End If
        
        gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ALTER COLUMN LeadEntryPointIsCorner INT"
    
        ' add SimpleEngraveFeed field if not there
        If Not mbln_DBFieldExists(r, "SimpleEngraveFeed") Then
            
            mint_UpdateDB = adoorUPDATE_DB_DOOR_PATHS_SIMPLE_ENGRAVE_FEED

            If gbln_SQLServer Then
                ' Update AWD_DOOR_PATHS to add ToolSidePartialReverse Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD SimpleEngraveFeed FLOAT"
            Else
                ' Update AWD_DOOR_PATHS to add ToolSidePartialReverse Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD SimpleEngraveFeed DOUBLE"
            End If
            
        End If
    
        ' add SimpleEngraveClearance field if not there
        If Not mbln_DBFieldExists(r, "SimpleEngraveClearance") Then
            
            mint_UpdateDB = adoorUPDATE_DB_DOOR_PATHS_SIMPLE_ENGRAVE_CLEARANCE

            If gbln_SQLServer Then
                ' Update AWD_DOOR_PATHS to add ToolSidePartialReverse Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD SimpleEngraveClearance FLOAT"
            Else
                ' Update AWD_DOOR_PATHS to add ToolSidePartialReverse Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_PATHS ADD SimpleEngraveClearance DOUBLE"
            End If
            
        End If
        
        If gbln_SQLServer Then
        
            g_DropConstraint "ALTER TABLE [dbo].[AD_DOOR_PATHS] DROP CONSTRAINT [SSMA_CC$AD_DOOR_PATHS$MachiningStyle$disallow_zero_length]"
        
        End If
    
    End If
    
    Call g_CloseRS(r)
    Set r = Nothing
  
    ' get all types from the database
    Set r = grst_GetAllTypes
    
    If Not (r Is Nothing) Then
                
        ' 19 mar 11 TFS#43510
        '
        ' add IgnoreOuterGeometry field if not there
        If Not mbln_DBFieldExists(r, "IgnoreOuterGeometry") Then
                
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_TYPES_IGNORE_OUTER_GEOMETRY
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_TYPES to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD IgnoreOuterGeometry BIT"
            Else
                ' Update AD_DOOR_TYPES to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD IgnoreOuterGeometry YESNO"
            End If
            
        End If
        
        ' add OversizeX field if not there
        If Not mbln_DBFieldExists(r, "OversizeX") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_TYPES_OVERSIZE_X
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD OversizeX FLOAT"
            Else
                ' Update AD_DOOR_PATHS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD OversizeX DOUBLE"
            End If
        
        End If
  
        ' add OversizeY field if not there
        If Not mbln_DBFieldExists(r, "OversizeY") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_TYPES_OVERSIZE_Y
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_PATHS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD OversizeY FLOAT"
            Else
                ' Update AD_DOOR_PATHS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD OversizeY DOUBLE"
            End If
        
        End If
    
        ' add PressID field if not there
        If Not mbln_DBFieldExists(r, "PressID") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_DOOR_TYPES_PRESS_ID
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_TYPES to add PressID Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD PressID INT"
            Else
                ' Update AD_DOOR_TYPES to add PressID Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD PressID LONG"
            End If
            
        End If
    
        ' add ColourID field if not there
        If Not mbln_DBFieldExists(r, "ColourID") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_DOOR_TYPES_COLOUR_ID
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_TYPES to add ColourID Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD ColourID INT"
            Else
                ' Update AD_DOOR_TYPES to add ColourID Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD ColourID LONG"
            End If
        End If
        
        
        If Not mbln_DBFieldExists(r, "ColourRotationMethod") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_DOOR_TYPES_ADD_ROTATION_METHOD
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_TYPES to add ColourRotationMethod Field
                gdb_CDM.Execute "ALTER TABLE dbo.AD_DOOR_TYPES ADD ColourRotationMethod INT NULL"
                ' Apply the constraint for the default value
                gdb_CDM.Execute "ALTER TABLE dbo.AD_DOOR_TYPES ADD CONSTRAINT DF_AD_DOOR_TYPES_ColourRotationMethod DEFAULT 0 FOR ColourRotationMethod"
            Else
                ' Update AD_DOOR_TYPES to add ColourRotationMethod Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD ColourRotationMethod INT DEFAULT 0"
            End If
        End If
        
        
        ' add HandleID field if not there
        If Not mbln_DBFieldExists(r, "HandleID") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_TYPES_ADD_HANDLE_ID
            
            If gbln_SQLServer Then
                ' Update AD_DOOR_TYPES to add HandleID Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD HandleID INT"
            Else
                ' Update AD_DOOR_TYPES to add HandleID Field
                gdb_CDM.Execute "ALTER TABLE AD_DOOR_TYPES ADD HandleID LONG"
            End If
        End If
        
    End If
    
    Call g_CloseRS(r)
    Set r = Nothing
    
    ' get all types from the database
    Set r = grst_GetAllOrderDetails
    
    If Not (r Is Nothing) Then
        
        ' 10 may 12 TFS#50399
        If Not mbln_DBFieldExists(r, "NestZoneID") Then
                
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_ADD_NEST_ZONE
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add NestZoneID Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD NestZoneID INT"
            Else
                ' Update AD_ORDER_DETAILS to add NestZoneID Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD NestZoneID LONG"
            End If
            
        End If
        
        ' add OversizeX field if not there
        If Not mbln_DBFieldExists(r, "OversizeX") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDERS_OVERSIZE_X
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD OversizeX FLOAT"
            Else
                ' Update AD_ORDER_DETAILS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD OversizeX DOUBLE"
            End If
            
        End If
  
        ' add OversizeY field if not there
        If Not mbln_DBFieldExists(r, "OversizeY") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDERS_OVERSIZE_Y
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD OversizeY FLOAT"
            Else
                ' Update AD_ORDER_DETAILS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD OversizeY DOUBLE"
            End If
        
        End If
            
        ' add the OriginalByPassNest field if not there
        If Not mbln_DBFieldExists(r, "OriginalByPassNest") Then
                
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORIGINAL_BYPASS
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD OriginalByPassNest BIT"
            Else
                ' Update AD_ORDER_DETAILS to add Engrave Corner Angle Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD OriginalByPassNest YESNO"
            End If
            
        End If
            
        ' add the ProductionComment field if not there
        If Not mbln_DBFieldExists(r, "ProductionComment") Then
                
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDERS_PRODUCTION_COMMENT
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add the ProductionComment Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ProductionComment VARCHAR(255)"
            Else
                ' Update AD_ORDER_DETAILS to add the ProductionComment Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ProductionComment MEMO"
            End If
            
        End If
      
        ' DOH! Active Reports cannot display a MEMO Field
        ' Change ProductionComment To type VARCHAR
    
        mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_PRODUCTION_COMMENT_CHANGE
        
        If Not gbln_SQLServer Then
            ' Change the data type for ProductionComment from Memo to Varchar
            gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ALTER COLUMN ProductionComment VARCHAR(255)"
        End If
    
        If r.Fields("StyleName").DefinedSize <> 255 Then
          
          ' Change the size of the StyleName field to 255 (was previously 50)
          mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_STYLE_NAME_FIELD_SIZE
      
          If gbln_SQLServer Then
              
              g_DropConstraint "ALTER TABLE [dbo].[AD_ORDER_DETAILS] DROP CONSTRAINT [SSMA_CC$AD_ORDER_DETAILS$StyleName$disallow_zero_length]"
              
              gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ALTER COLUMN StyleName VARCHAR(255)"
              
              gdb_CDM.Execute "ALTER TABLE [dbo].[AD_ORDER_DETAILS]  WITH CHECK ADD  CONSTRAINT [SSMA_CC$AD_ORDER_DETAILS$StyleName$disallow_zero_length] CHECK  ((len([StyleName])>(0)))"
              gdb_CDM.Execute "ALTER TABLE [dbo].[AD_ORDER_DETAILS] CHECK CONSTRAINT [SSMA_CC$AD_ORDER_DETAILS$StyleName$disallow_zero_length]"
          
          Else
              gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ALTER COLUMN StyleName VARCHAR(255)"
          End If
        
        End If
        
        ' add the CustomField1 field if not there
        If Not mbln_DBFieldExists(r, "CustomField1") Then
                
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_CUSTOM_FIELD1
            
            ' Update AD_ORDER_DETAILS to add the CustomField1 Field
            gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD CustomField1 VARCHAR(255)"
            
        End If
        
        ' add the CustomField2 field if not there
        If Not mbln_DBFieldExists(r, "CustomField2") Then
                
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_CUSTOM_FIELD2
            
            ' Update AD_ORDER_DETAILS to add the CustomField2 Field
            gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD CustomField2 VARCHAR(255)"
            
        End If
    
        ' add PressID field if not there
        If Not mbln_DBFieldExists(r, "PressID") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_PRESS_ID
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add PressID Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD PressID INT"
            Else
                ' Update AD_ORDER_DETAILS to add PressID Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD PressID LONG"
            End If
        End If
    
        ' add ColourID field if not there
        If Not mbln_DBFieldExists(r, "ColourID") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_COLOUR_ID
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add ColourID Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ColourID INT"
            Else
                ' Update AD_ORDER_DETAILS to add ColourID Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ColourID LONG"
            End If
        End If
        
        ' add PostProcessor field if not there
        If Not mbln_DBFieldExists(r, "PostProcessor") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_POST_PROCESSOR
            
            ' Update AD_ORDER_DETAILS to add ColourID Field
            gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD PostProcessor VARCHAR(255)"
            
        End If
        
        ' add ReverseMachiningFilename field if not there
        If Not mbln_DBFieldExists(r, "ReverseMachiningFilename") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_REVERSE_MACHINING_FILENAME
            
            ' Update AD_ORDER_DETAILS to add ReverseMachiningFilename Field
            gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ReverseMachiningFilename VARCHAR(255)"
            
        End If
        
        ' add Shop Floor Data Capture (SFDC_Approve) field if not there
        If Not mbln_DBFieldExists(r, "SFDC_Approve") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_SFDC_APPROVE
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add SFDC_Approve Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD SFDC_Approve INT"
            Else
                ' Update AD_ORDER_DETAILS to add SFDC_Approve Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD SFDC_Approve LONG"
            End If
        
        End If
    
        ' add Shop Floor Data Capture (SFDC_Reject) field if not there
        If Not mbln_DBFieldExists(r, "SFDC_Reject") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_SFDC_REJECT
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add SFDC_Approve Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD SFDC_Reject INT"
            Else
                ' Update AD_ORDER_DETAILS to add SFDC_Approve Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD SFDC_Reject LONG"
            End If
            
        End If
    
        ' add Shop Floor Data Capture (Approve Stage 2) field if not there
        If Not mbln_DBFieldExists(r, "SFDC_Approve_Stage_2") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_SFDC_APPROVE_STAGE_2
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add SFDC_Approve_Stage_2 Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD SFDC_Approve_Stage_2 INT"
            Else
                ' Update AD_ORDER_DETAILS to add SFDC_Approve_Stage_2 Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD SFDC_Approve_Stage_2 LONG"
            End If
            
        End If
    
        ' add Shop Floor Data Capture (Approve Stage 3) field if not there
        If Not mbln_DBFieldExists(r, "SFDC_Approve_Stage_3") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_SFDC_APPROVE_STAGE_3
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add SFDC_Approve_Stage_3 Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD SFDC_Approve_Stage_3 INT"
            Else
                ' Update AD_ORDER_DETAILS to add SFDC_Approve_Stage_3 Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD SFDC_Approve_Stage_3 LONG"
            End If
        End If
        
        ' add ComponentGrouping field if not there
        If Not mbln_DBFieldExists(r, "ComponentGrouping") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_COMPONENT_GROUPING
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add ComponentGrouping Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ComponentGrouping INT"
            Else
                ' Update AD_ORDER_DETAILS to add ComponentGrouping Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ComponentGrouping LONG"
            End If
        
        End If
    
        ' add ReworkQuantity field if not there
        If Not mbln_DBFieldExists(r, "ReworkQuantity") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_SFDC_REWORK_QUANTITY
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add ReworkQuantity Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ReworkQuantity INT"
            Else
                ' Update AD_ORDER_DETAILS to add ReworkQuantity Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ReworkQuantity LONG"
            End If
            
        End If
    
        If Not mbln_DBFieldExists(r, "ColourRotationMethod") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_ADD_ROTATION_METHOD
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add ColourRotationMethod Field
                gdb_CDM.Execute "ALTER TABLE dbo.AD_ORDER_DETAILS ADD ColourRotationMethod INT NULL"
                ' Apply the constraint for the default value
                gdb_CDM.Execute "ALTER TABLE dbo.AD_ORDER_DETAILS ADD CONSTRAINT DF_AD_ORDER_DETAILS_ColourRotationMethod DEFAULT 0 FOR ColourRotationMethod"
            Else
                ' Update AD_ORDER_DETAILS to add ColourRotationMethod Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD ColourRotationMethod INT DEFAULT 0"
            End If
        End If
    
        ' add HandleID field if not there
        If Not mbln_DBFieldExists(r, "HandleID") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_ADD_HANDLE_ID
            
            If gbln_SQLServer Then
                ' Update AD_ORDER_DETAILS to add HandleID Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD HandleID INT"
            Else
                ' Update AD_ORDER_DETAILS to add HandleID Field
                gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS ADD HandleID LONG"
            End If
        End If
    
    End If
    
    Call g_CloseRS(r)
    Set r = Nothing
    
    ' get all report data from the database
    Set r = grst_GetAllReportData
    
    If Not (r Is Nothing) Then
        
        ' add the ProductionComment field if not there
        If Not mbln_DBFieldExists(r, "ProductionComment") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_PRODUCTION_COMMENT
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the ProductionComment Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD ProductionComment VARCHAR(255)"
            Else
                ' Update AD_REPORT_DATA to add the ProductionComment Field
                ' Changed from Memo field type 8th Aug 13
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD ProductionComment VARCHAR(255)"
            End If
        End If
        
        Call g_CloseRS(r)
        Set r = Nothing
    
    End If
    
    ' get all report data from the database
    Set r = grst_GetAllReportData
    
    If Not (r Is Nothing) Then
        
        ' DOH! Active Reports cannot display a MEMO Field
        
        If Not gbln_SQLServer Then
            ' Change ProductionComment To type VARCHAR
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_PRODUCTION_COMMENT_CHANGE
            
            If r.Fields!ProductionComment.Type <> adVarWChar Then
              ' Change the data type for ProductionComment from Memo to Varchar
              gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ALTER COLUMN ProductionComment VARCHAR(255)"
            End If
        End If
        
        mint_UpdateDB = adoorUPDATE_DB_REPORT_DATA_PARTITEMNUMBER_CHANGE
                
        ' Update AD_REPORT_DATA and change the column PartItemNumber from Text to Number
        If r.Fields!PartItemNumber.Type <> adInteger Then
        
          If gbln_SQLServer Then
              g_DropConstraint "ALTER TABLE [dbo].[AD_REPORT_DATA] DROP CONSTRAINT [DF__AD_REPORT__PartI__797309D9]"
              gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ALTER COLUMN PartItemNumber INT"
          Else
              gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ALTER COLUMN PartItemNumber INT DEFAULT 0"
          End If
          
        End If
        
        ' add the PressName field if not there
        If Not mbln_DBFieldExists(r, "PressName") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_NAME
            
            ' Update AD_REPORT_DATA to add the PressName Field
            gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressName VARCHAR(255)"
            
        End If
    
        ' add the FoilColour field if not there
        If Not mbln_DBFieldExists(r, "FoilColour") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_FOIL_COLOUR
            
            ' Update AD_REPORT_DATA to add the FoilColour Field
            gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD FoilColour VARCHAR(255)"
            
        End If
    
        ' add the PressSheetName field if not there
        If Not mbln_DBFieldExists(r, "PressSheetName") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_SHEET_NAME
            
            ' Update AD_REPORT_DATA to add the PressSheetName Field
            gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressSheetName VARCHAR(50)"
            
        End If
    
        ' add the PressItemNumber field if not there
        If Not mbln_DBFieldExists(r, "PressItemNumber") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_ITEM_NUMBER
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the PressItemNumber Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressItemNumber INT"
            Else
                ' Update AD_REPORT_DATA to add the PressItemNumber Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressItemNumber LONG"
            End If
            
        End If
    
        ' add the PressQuantityOnSheet field if not there
        If Not mbln_DBFieldExists(r, "PressQuantityOnSheet") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_QUANTITY_ON_SHEET
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the PressQuantityOnSheet Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressQuantityOnSheet INT"
            Else
                ' Update AD_REPORT_DATA to add the PressQuantityOnSheet Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressQuantityOnSheet LONG"
            End If
            
        End If
    
        ' add the PressPathToEMF field if not there
        If Not mbln_DBFieldExists(r, "PressPathToEMF") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_PATH_TO_EMF
            
            ' Update AD_REPORT_DATA to add the PressPathToEMF Field
            gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressPathToEMF VARCHAR(255)"
            
        End If
    
        ' add the PressQuantityThisSheet field if not there
        If Not mbln_DBFieldExists(r, "PressQuantityThisSheet") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_QUANTITY_THIS_SHEET
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the PressQuantityThisSheet Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressQuantityThisSheet INT"
            Else
                ' Update AD_REPORT_DATA to add the PressQuantityThisSheet Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressQuantityThisSheet LONG"
            End If
        
        End If
    
        ' add the PressDoorImage field if not there
        If Not mbln_DBFieldExists(r, "PressDoorImage") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_DOOR_IMAGE
            
            ' Update AD_REPORT_DATA to add the PressDoorImage Field
            gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressDoorImage VARCHAR(255)"
            
        End If
    
        ' add the PressSheetIdentifier field if not there
        If Not mbln_DBFieldExists(r, "PressSheetIdentifier") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_SHEET_IDENTIFIER
            
            ' Update AD_REPORT_DATA to add the PressSheetIdentifier Field
            gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressSheetIdentifier VARCHAR(50)"
            
        End If
    
        ' add the PressDoorCounter field if not there
        If Not mbln_DBFieldExists(r, "PressDoorCounter") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_SHEET_IDENTIFIER
            
            ' Update AD_REPORT_DATA to add the PressSheetIdentifier Field
            gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressDoorCounter INT DEFAULT 0"
            
        End If
    
        ' add the LabelPrinted field if not there
        If Not mbln_DBFieldExists(r, "LabelPrinted") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_LABEL_PRINTED
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the LabelPrinted Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD LabelPrinted INT"
            Else
                ' Update AD_REPORT_DATA to add the LabelPrinted Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD LabelPrinted LONG"
            End If
            
        End If
    
        ' add the NestPartPositionLeft field if not there
        If Not mbln_DBFieldExists(r, "NestPartPositionLeft") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_NEST_PART_LEFT
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the NestPartPositionLeft Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD NestPartPositionLeft FLOAT"
            Else
                ' Update AD_REPORT_DATA to add the NestPartPositionLeft Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD NestPartPositionLeft DOUBLE"
            End If
            
        End If
    
        ' add the NestPartPositionRight field if not there
        If Not mbln_DBFieldExists(r, "NestPartPositionRight") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_NEST_PART_RIGHT
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the NestPartPositionRight Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD NestPartPositionRight FLOAT"
            Else
                ' Update AD_REPORT_DATA to add the NestPartPositionRight Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD NestPartPositionRight DOUBLE"
            End If
            
        End If
    
        ' add the NestPartPositionTop field if not there
        If Not mbln_DBFieldExists(r, "NestPartPositionTop") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_NEST_PART_TOP
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the NestPartPositionTop Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD NestPartPositionTop FLOAT"
            Else
                ' Update AD_REPORT_DATA to add the NestPartPositionTop Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD NestPartPositionTop DOUBLE"
            End If
            
        End If
    
        ' add the NestPartPositionBottom field if not there
        If Not mbln_DBFieldExists(r, "NestPartPositionBottom") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_NEST_PART_BOTTOM
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the NestPartPositionBottom Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD NestPartPositionBottom FLOAT"
            Else
                ' Update AD_REPORT_DATA to add the NestPartPositionBottom Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD NestPartPositionBottom DOUBLE"
            End If
            
        End If
    
        ' add the PressSheetNumber field if not there
        If Not mbln_DBFieldExists(r, "PressSheetNumber") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PRESS_SHEET_NUMBER
            
            If gbln_SQLServer Then
                ' Update AD_REPORT_DATA to add the PressSheetNumber Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressSheetNumber INT"
            Else
                ' Update AD_REPORT_DATA to add the PressSheetNumber Field
                gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PressSheetNumber LONG"
            End If
            
        End If
        
        ' add the PathToPressNestARD field if not there
        If Not mbln_DBFieldExists(r, "PathToPressNestARD") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_REPORTS_ADD_PATH_TO_PRESS_NEST_ARD
            
            ' Update AD_REPORT_DATA to add the PathToPressNestARD Field
            gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA ADD PathToPressNestARD VARCHAR(255)"
            
        End If
        
    End If
    
    Call g_CloseRS(r)
    Set r = Nothing
    
    ' get all types from the database
    Set r = grst_GetMaterials
    If Not (r Is Nothing) Then
        
        ' add Leave Edge Gap field if not there
        If Not mbln_DBFieldExists(r, "LeaveEdgeGapUncut") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_MATERIALS_LEAVE_EDGE_GAP
            
            If gbln_SQLServer Then
                ' Update AD_MATERIALS to add Leave Edge Gap Uncut Field
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD LeaveEdgeGapUncut BIT"
            Else
                ' Update AD_MATERIALS to add Leave Edge Gap Uncut Field
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD LeaveEdgeGapUncut YESNO"
            End If
        End If
  
        ' add NumberComponentsBySize field if not there
        If Not mbln_DBFieldExists(r, "NumberComponentsBySize") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_MATERIALS_NUMBER_COMPONENTS_BY_SIZE
            
            If gbln_SQLServer Then
                ' Update AD_MATERIALS to add Biggest Parts First Field
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD NumberComponentsBySize BIT"
            Else
                ' Update AD_MATERIALS to add Biggest Parts First Field
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD NumberComponentsBySize YESNO"
            End If
        End If
    
        ' add Suppress Final Sort field if not there
        If Not mbln_DBFieldExists(r, "SuppressFinalSort") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_MATERIALS_SUPPRESS_FINAL_SORT
            
            If gbln_SQLServer Then
                ' Update AD_MATERIALS to add Suppress Nesting Final Sort Field
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD SuppressFinalSort BIT"
            Else
                ' Update AD_MATERIALS to add Suppress Nesting Final Sort Field
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD SuppressFinalSort YESNO"
            End If
            
        End If
    
        ' add NestingScreenUpdate field if not there
        If Not mbln_DBFieldExists(r, "NestingScreenUpdate") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_MATERIALS_NESTING_SCREEN_UPDATE
            
            If gbln_SQLServer Then
                ' Update AD_MATERIALS to add NestingScreenUpdate Field
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD NestingScreenUpdate BIT"
            Else
                ' Update AD_MATERIALS to add NestingScreenUpdate Field
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD NestingScreenUpdate YESNO"
            End If
        End If
    
        ' 04/20/06 - rg
        '
        ' --> add OnionSkin fields if not there
        '
        If Not mbln_DBFieldExists(r, "OnionSkin") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ONION_SKIN
                If gbln_SQLServer Then
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkin BIT"
                Else
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkin YESNO"
                End If
        End If
        
        If Not mbln_DBFieldExists(r, "OnionSkinMinXY") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ONION_SKIN_MINXY
                If gbln_SQLServer Then
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkinMinXY FLOAT"
                Else
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkinMinXY DOUBLE"
                End If
        End If
        
        If Not mbln_DBFieldExists(r, "OnionSkinMinArea") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ONION_SKIN_MINAREA
                If gbln_SQLServer Then
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkinMinArea FLOAT"
                Else
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkinMinArea DOUBLE"
                End If
        End If
        
        If Not mbln_DBFieldExists(r, "OnionSkinThickness") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ONION_SKIN_THICKNESS
                If gbln_SQLServer Then
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkinThickness FLOAT"
                Else
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkinThickness DOUBLE"
                End If
        End If
        
        If Not mbln_DBFieldExists(r, "OnionSkinCutOrder") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ONION_SKIN_CUTORDER
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkinCutOrder INT"
        End If
        
        If Not mbln_DBFieldExists(r, "OnionSkinApplyToInside") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ONION_SKIN_APPLYTOINSIDE
                If gbln_SQLServer Then
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkinApplyToInside BIT"
                Else
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD OnionSkinApplyToInside YESNO"
                End If
        End If
        '
        ' <--
    
        If Not mbln_DBFieldExists(r, "ProcessWaste") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_PROCESS_WASTE
                If gbln_SQLServer Then
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD ProcessWaste BIT"
                Else
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD ProcessWaste YESNO"
                End If
        End If
    
        If Not mbln_DBFieldExists(r, "ProcessWasteMCStyle") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_PROCESS_WASTE_MC_STYLE
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD ProcessWasteMCStyle VARCHAR(255)"
        End If
    
        If Not mbln_DBFieldExists(r, "ProcessWasteDepthOfCut") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_PROCESS_WASTE_DEPTH_OF_CUT
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD ProcessWasteDepthOfCut INT"
        End If
    
        If Not mbln_DBFieldExists(r, "ProcessWasteFinalSheetScrap") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_PROCESS_WASTE_DEPTH_OF_CUT
                If gbln_SQLServer Then
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD ProcessWasteFinalSheetScrap FLOAT"
                Else
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD ProcessWasteFinalSheetScrap DOUBLE"
                End If
        End If
    
        If Not mbln_DBFieldExists(r, "ProcessWasteCutTowardsComponents") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_PROCESS_WASTE_CUT_TOWARDS_COMPONENTS
                If gbln_SQLServer Then
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD ProcessWasteCutTowardsComponents BIT"
                Else
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD ProcessWasteCutTowardsComponents YESNO"
                End If
        End If
    
        If Not mbln_DBFieldExists(r, "InsertFillerFile2") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_FILLER_FILE_2
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD InsertFillerFile2 VARCHAR(255)"
        End If
    
        If Not mbln_DBFieldExists(r, "InsertFillerFile3") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_FILLER_FILE_3
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD InsertFillerFile3 VARCHAR(255)"
        End If
    
        If Not mbln_DBFieldExists(r, "ProcessWasteStrategy") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ADD_PROCESS_WASTE_STRATEGY
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD ProcessWasteStrategy INT"
        End If
    
        If Not mbln_DBFieldExists(r, "HorizontalCutSpacing") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ADD_HORIZONTAL_CUT_SPACING
                If gbln_SQLServer Then
                  gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD HorizontalCutSpacing FLOAT"
                Else
                  gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD HorizontalCutSpacing DOUBLE"
                End If
        End If
    
        If Not mbln_DBFieldExists(r, "VerticalCutSpacing") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ADD_VERTICAL_CUT_SPACING
                If gbln_SQLServer Then
                  gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD VerticalCutSpacing FLOAT"
                Else
                  gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD VerticalCutSpacing DOUBLE"
                End If
        End If
    
        If Not mbln_DBFieldExists(r, "PackFinalSheetComponentsToLHS") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ADD_PACK_FINAL_SHEET_TO_LHS
                If gbln_SQLServer Then
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD PackFinalSheetComponentsToLHS BIT"
                Else
                    gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD PackFinalSheetComponentsToLHS YESNO"
                End If
        End If
    
        If Not mbln_DBFieldExists(r, "TimePerSheet") Then
                mint_UpdateDB = adoorUPDATE_DB_MATERIALS_ADD_TIME_PER_SHEET
                gdb_CDM.Execute "ALTER TABLE AD_MATERIALS ADD TimePerSheet INT"
        End If
    
    End If
    
    Call g_CloseRS(r)
    Set r = Nothing
    
    ' get all types from the database
    Set r = grst_GetTables
    If Not (r Is Nothing) Then
        
        If Not mbln_DBTableExists(r, "AD_USER_STYLES") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_USER_STYLES_TABLE
            
            If gbln_SQLServer Then
                
                gdb_CDM.Execute "CREATE TABLE AD_USER_STYLES (UserStyleID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, FullFileName VARCHAR(255) NOT NULL,VBAProjectName VARCHAR(255) NOT NULL)"
            
            Else
                gdb_CDM.Execute "CREATE TABLE AD_USER_STYLES (UserStyleID COUNTER PRIMARY KEY, FullFileName VARCHAR(255) NOT NULL,VBAProjectName VARCHAR(255) NOT NULL)"
            End If
            
        End If
        
        If Not mbln_DBTableExists(r, "AD_DOOR_TREE") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_DOOR_TREE_TABLE
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_DOOR_TREE (ID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, NodeIndex INT, NodeParentIndex INT, NodeText VARCHAR(100), NodeKey VARCHAR(255), NodeDataKey VARCHAR(10), NodeImage VARCHAR(50), NodeTag VARCHAR(10), NodeTagVariant VARCHAR(100))"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_DOOR_TREE (ID COUNTER PRIMARY KEY, NodeIndex LONG, NodeParentIndex LONG, NodeText VARCHAR(100), NodeKey VARCHAR(255), NodeDataKey VARCHAR(10), NodeImage VARCHAR(50), NodeTag VARCHAR(10), NodeTagVariant VARCHAR(100))"
            End If
        
        End If
    
        If Not mbln_DBTableExists(r, "AD_ORDER_GRID") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_ORDER_GRID_TABLE
            
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_ORDER_GRID (ID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, Jobname VARCHAR(255), PO VARCHAR(255), OrderDate VARCHAR(50), DueDate VARCHAR(50), ProcessedDate VARCHAR(50), TotalParts INT, OrderID INT)"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_ORDER_GRID (ID COUNTER PRIMARY KEY, Jobname VARCHAR(255), PO VARCHAR(255), OrderDate VARCHAR(50), DueDate VARCHAR(50), ProcessedDate VARCHAR(50), TotalParts LONG, OrderID LONG)"
            End If
        End If
    
        If Not mbln_DBTableExists(r, "AD_SETTINGS") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_SETTINGS_TABLE
            
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_SETTINGS (NCSeqNum INT)"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_SETTINGS (NCSeqNum LONG)"
            End If
            
            Dim l As Long
            
            ' set the initial value when the table has been created
            gdb_CDM.Execute "INSERT INTO AD_SETTINGS(NCSeqNum) VALUES(1)", l
        
        End If
    
        If Not mbln_DBTableExists(r, "AD_PRESSES") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_PRESSES_TABLE
            
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_PRESSES (PressID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, PressName VARCHAR(255), PressWidth FLOAT, PressLength FLOAT, DefaultPress BIT, MinGapBetweenPaths FLOAT, GapAtSheetEdge FLOAT, PartRotation INT, NumberComponentsBySize BIT)"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_PRESSES (PressID COUNTER PRIMARY KEY, PressName VARCHAR(255), PressWidth DOUBLE, PressLength DOUBLE, DefaultPress YESNO, MinGapBetweenPaths DOUBLE, GapAtSheetEdge DOUBLE, PartRotation INT, NumberComponentsBySize YESNO)"
            End If
            
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_POPULATE_PRESSES_TABLE
        
            gdb_CDM.Execute _
              "INSERT INTO AD_PRESSES(PressName, PressWidth, PressLength, DefaultPress, MinGapBetweenPaths, " & _
              "GapAtSheetEdge, PartRotation, NumberComponentsBySize)" & _
              "VALUES ('Example Press',1220,2440,-1,5,2,0,0)"
        
        End If
        
        If Not mbln_DBTableExists(r, "AD_MATERIAL_COLOURS") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_COLOURS_TABLE
            
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_MATERIAL_COLOURS (ColourID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, ColourName VARCHAR(255))"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_MATERIAL_COLOURS (ColourID COUNTER PRIMARY KEY, ColourName VARCHAR(255))"
            End If
        End If
        
        If Not mbln_DBTableExists(r, "AD_TOOL_ORDER") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_TOOL_ORDER_TABLE
            
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_TOOL_ORDER (ToolOrderID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, ToolName VARCHAR(255), ToolNumber INT, ToolOffsetNumber INT, ToolSeqNum INT)"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_TOOL_ORDER (ToolOrderID COUNTER PRIMARY KEY, ToolName VARCHAR(255), ToolNumber INT, ToolOffsetNumber INT, ToolSeqNum INT)"
            End If
        
        End If
        
        If Not mbln_DBTableExists(r, "AD_NEST_ZONES") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_NEST_ZONES_TABLE
            
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_NEST_ZONES (ZoneID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, ZoneName VARCHAR(255), ZoneNumber INT, ZoneStartPointX FLOAT, ZoneStartPointY FLOAT, ZoneWidth FLOAT, ZoneHeight FLOAT, ZoneAssignMethod INT, ZonePartDimension FLOAT, ZonePartArea FLOAT)"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_NEST_ZONES (ZoneID COUNTER PRIMARY KEY, ZoneName VARCHAR(255), ZoneNumber INT, ZoneStartPointX DOUBLE, ZoneStartPointY DOUBLE, ZoneWidth DOUBLE, ZoneHeight DOUBLE, ZoneAssignMethod INT, ZonePartDimension DOUBLE, ZonePartArea DOUBLE)"
            End If
            
        End If
                
        If Not mbln_DBTableExists(r, "AD_HANDLE_DRILLING") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_HANDLE_DRILLING_TABLE
            
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_HANDLE_DRILLING (HandleID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, HandleName VARCHAR(255), DatumLocation INT, DatumLocationAdditionalOffsetX FLOAT, DatumLocationAdditionalOffsetY FLOAT, NumberOfHoles INT, HoleOrientation INT, HoleSpacing FLOAT, HoleSize FLOAT, MachiningStyle VARCHAR(255))"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_HANDLE_DRILLING (HandleID COUNTER PRIMARY KEY, HandleName VARCHAR(255), DatumLocation INT, DatumLocationAdditionalOffsetX DOUBLE, DatumLocationAdditionalOffsetY DOUBLE, NumberOfHoles INT, HoleOrientation INT, HoleSpacing DOUBLE, HoleSize DOUBLE, MachiningStyle VARCHAR(255))"
            End If
            
        End If
        
    End If
        
    Call g_CloseRS(r)
    Set r = Nothing
    
    Set r = grst_GetPresses
    If Not (r Is Nothing) Then
        
        ' add UseTrueShape field if not there
        If Not mbln_DBFieldExists(r, "UseTrueShape") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_PRESSES_ADD_USE_TRUE_SHAPE
            
            If gbln_SQLServer Then
                ' Update AD_PRESSES to add UseTrueShape Field
                gdb_CDM.Execute "ALTER TABLE AD_PRESSES ADD UseTrueShape BIT"
            Else
                ' Update AD_PRESSES to add UseTrueShape Field
                gdb_CDM.Execute "ALTER TABLE AD_PRESSES ADD UseTrueShape YESNO"
            End If
        End If
    
        ' add TrueShapePackTo field if not there
        If Not mbln_DBFieldExists(r, "TrueShapePackTo") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_PRESSES_ADD_TRUE_SHAPE_PACK_TO
            
            ' Update AD_PRESSES to add TrueShapePackTo Field
            gdb_CDM.Execute "ALTER TABLE AD_PRESSES ADD TrueShapePackTo INT"
            
        End If
    
        ' add RectPackTo field if not there
        If Not mbln_DBFieldExists(r, "RectPackTo") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_PRESSES_ADD_RECT_PACK_TO
            
            ' Update AD_PRESSES to add RectPackTo Field
            gdb_CDM.Execute "ALTER TABLE AD_PRESSES ADD RectPackTo INT"
            
        End If
    
    End If
        
    Call g_CloseRS(r)
    Set r = Nothing
    
    Set r = grst_GetColours
    If Not (r Is Nothing) Then
        
        ' add ColourRotationMethod field if not there
        If Not mbln_DBFieldExists(r, "ColourRotationMethod") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_COLOURS_ADD_ROTATION_METHOD
            
            If gbln_SQLServer Then
                ' Update AD_MATERIAL_COLOURS to add ColourRotationMethod Field
                gdb_CDM.Execute "ALTER TABLE dbo.AD_MATERIAL_COLOURS ADD ColourRotationMethod INT NULL"
                ' Apply the default value constraint
                gdb_CDM.Execute "ALTER TABLE dbo.AD_MATERIAL_COLOURS ADD CONSTRAINT DF_AD_MATERIAL_COLOURS_ColourRotationMethod DEFAULT 0 FOR ColourRotationMethod"
            Else
                ' Update AD_MATERIAL_COLOURS to add ColourRotationMethod Field
                gdb_CDM.Execute "ALTER TABLE AD_MATERIAL_COLOURS ADD ColourRotationMethod INT DEFAULT 0"
            End If
        End If
    End If
    
    Set r = grst_GetTables
    
    If Not r Is Nothing Then
    
        If Not mbln_DBTableExists(r, "AD_RULES") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_RULES_TABLE
            
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_RULES (RuleID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, OperatorID INT, RuleName VARCHAR(255), TestVariableName INT, TestVariableValue FLOAT, ResultRouterPost VARCHAR(255), ResultRouter INT, RuleText VARCHAR(255))"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_RULES (RuleID COUNTER PRIMARY KEY, OperatorID INT, RuleName VARCHAR(255), TestVariableName INT, TestVariableValue DOUBLE, ResultRouterPost VARCHAR(255), ResultRouter INT, RuleText VARCHAR(255))"
            End If
            
        End If
        
        If Not mbln_DBTableExists(r, "AD_OPERATORS") Then
        
            ' up the return flag
            mint_UpdateDB = adoorUPDATE_DB_ADD_OPERATORS_TABLE
            
            If gbln_SQLServer Then
                gdb_CDM.Execute "CREATE TABLE AD_OPERATORS (OperatorID INT NOT NULL IDENTITY (1, 1) PRIMARY KEY, Operator VARCHAR(5), OperatorDescription VARCHAR(100))"
            Else
                gdb_CDM.Execute "CREATE TABLE AD_OPERATORS (OperatorID COUNTER PRIMARY KEY, Operator VARCHAR(5), OperatorDescription VARCHAR(100))"
            End If
        
            mint_UpdateDB = adoorUPDATE_DB_POPULATE_OPERATORS_TABLE
            
            ' Populate the AD_Operators Table
            If Not mbln_InsertOperators Then
              GoTo Controlled_Exit
            End If
        
        End If
        
        mint_UpdateDB = adoorUPDATE_DB_POPULATE_OPERATORS_TABLE
        
        If mbln_TestUpdateOperators Then
          If Not mbln_UpdateOperatorDescriptions Then
            GoTo Controlled_Exit
          End If
        End If
        
    End If
    
    Call g_CloseRS(r)
    Set r = Nothing
    
    ' Add Zones Table
    
    
    
'    **************************************************************************************************
'    Removed - 080110
'    This can fail on a database with a large amount of information in the AD_REPORT_DATA table
'    This was only added for SQL Server Migration compatibility (the code below appears
'    to work, but the tables still have to be opened in Access anyway (and saved) for the SQL Migration
'    tool to work
'    **************************************************************************************************
'    ' Only do the unicode stuff if using Microsoft Access
'    If Not gbln_SQLServer Then
'
'        ' Modify the field 'ProductionComment' to set Unicode Compression to YES
'        ' This enables seemless migration to SQL Server
'        mint_UpdateDB = adoorUPDATE_DB_ORDER_DETAILS_PRODCOMMENT_UNICODE_COMP
'        gdb_CDM.Execute "ALTER TABLE AD_ORDER_DETAILS " & _
'          "ALTER COLUMN ProductionComment TEXT(255) WITH COMPRESSION"
'
'        ' Modify the field 'ProductionComment' to set Unicode Compression to YES
'        ' This enables seemless migration to SQL Server
'        mint_UpdateDB = adoorUPDATE_DB_REPORT_DATA_PRODCOMMENT_UNICODE_COMP
'        gdb_CDM.Execute "ALTER TABLE AD_REPORT_DATA " & _
'          "ALTER COLUMN ProductionComment TEXT(255) WITH COMPRESSION"
'
'        ' Modify the field 'ProcessWasteMCStyle' to set Unicode Compression to YES
'        ' This enables seemless migration to SQL Server
'        mint_UpdateDB = adoorUPDATE_DB_MATERIALS_PROCESSWASTEMCSTYLE_UNICODE_COMP
'        gdb_CDM.Execute "ALTER TABLE AD_MATERIALS " & _
'          "ALTER COLUMN ProcessWasteMCStyle TEXT(255) WITH COMPRESSION"
'
'        ' Modify the field 'InsertFillerFile2' to set Unicode Compression to YES
'        ' This enables seemless migration to SQL Server
'        mint_UpdateDB = adoorUPDATE_DB_MATERIALS_INSERTFILLER2_UNICODE_COMP
'        gdb_CDM.Execute "ALTER TABLE AD_MATERIALS " & _
'          "ALTER COLUMN InsertFillerFile2 TEXT(255) WITH COMPRESSION"
'
'        ' Modify the field 'InsertFillerFile3' to set Unicode Compression to YES
'        ' This enables seemless migration to SQL Server
'        mint_UpdateDB = adoorUPDATE_DB_MATERIALS_INSERTFILLER3_UNICODE_COMP
'        gdb_CDM.Execute "ALTER TABLE AD_MATERIALS " & _
'          "ALTER COLUMN InsertFillerFile3 TEXT(255) WITH COMPRESSION"
'
'    End If
    
    ' set success return flag
    mint_UpdateDB = adoorUPDATE_DB_SUCCESS
    
Controlled_Exit:
    
    Call g_CloseRS(r)
    Set r = Nothing
  
Exit Function
  
EH:
    
    Debug.Print Err.Description & ": " & Err.Source
    
    Select Case Err.Number
      Case -2147217900
    
      Case Else
        'MsgBox Err.Number & " to be added"
      
    End Select
    
    
    Resume Controlled_Exit
    
End Function

Private Function mbln_DBFieldExists(ByVal rRecordset As ADODB.Recordset, ByVal sFieldName As String) As Boolean ' 07/25/03 - rg
        
    Dim oField                  As ADODB.Field
    Dim oFields                 As ADODB.Fields
    
On Error GoTo EH:
    
    ' assume not
    mbln_DBFieldExists = False
    
    ' 16 feb 12 - rg
    '
    If (rRecordset Is Nothing) Then GoTo Controlled_Exit
    
    ' get the fields from the passed recordset
    Set oFields = rRecordset.Fields
    
    ' loop thru all the fiels looking for passed field name - bail if found
    For Each oField In oFields
               
        If (StrComp(oField.Name, sFieldName, vbTextCompare) = 0) Then
            mbln_DBFieldExists = True
            Exit For
        End If
        
    Next oField

Controlled_Exit:
        
    Set oField = Nothing
    Set oFields = Nothing

Exit Function

EH:
    
    MsgBox Err.Description, vbExclamation, DEF_REG_APP_NAME
    mbln_DBFieldExists = False
    Resume Controlled_Exit
    
End Function

Private Function mbln_DBTableExists(ByVal rRecordset As ADODB.Recordset, ByVal sTableName As String) As Boolean
        
    
On Error GoTo EH:
    
    ' assume not
    mbln_DBTableExists = False
    
    rRecordset.MoveFirst
    
    ' loop thru all the tables looking for passed table name - bail if found
    Do While Not rRecordset.EOF
        
        If (StrComp(rRecordset.Fields("Table_Name"), sTableName, vbTextCompare) = 0) Then
            mbln_DBTableExists = True
            Exit Do
        End If
        
        rRecordset.MoveNext
    Loop

Controlled_Exit:
        

Exit Function

EH:
    
    MsgBox Err.Description, vbExclamation, DEF_REG_APP_NAME
    mbln_DBTableExists = False
    Resume Controlled_Exit
    
End Function


Function OnUpdatem_OrderToolPaths()
  If App.ActiveDrawing.GetToolPathCount > 0 Then
    OnUpdatem_OrderToolPaths = 1
  Else
    OnUpdatem_OrderToolPaths = 0
  End If
End Function

Function OnUpdatem_Processing()
'
'  On Error GoTo EH:
'
'  If m_AlphaDOORRunning Then
'      OnUpdatem_Processing = 0
'  Else
'      OnUpdatem_Processing = 1
'  End If
'
'Exit Function
'
'EH:
  OnUpdatem_Processing = 1
End Function
 

Function m_Processing()

    Dim oCDM                    As Object
    Dim oSplash                 As Object
    Dim clsOptions              As COptions
    
On Error GoTo m_Processing_Error
    
    Set clsOptions = New COptions
    strCTX = clsOptions.CTXFile
        
    '..save the path to this macro in the registry
    SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "MacroPath", Frame.PathOfThisAddin
    
    With App
    
        If .ActiveDrawing.Geometries.Count > 0 Then
            
            '..continue?
            If MsgBox(.Frame.ReadTextFile(.Frame.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 33), _
                      vbYesNo + vbQuestion, DEF_PROJECT_NAME) = vbNo Then GoTo Controlled_Exit
            
        
        End If
    
        '..start new drawing
        .New
    
    End With

    '..check for registered user
    If Not mbln_OKtoRun Then GoTo Controlled_Exit
            
    If clsOptions.BackupDatabase And clsOptions.BackupDatabaseStartUp Then
        g_CloseDBConnection
        DoEvents
        If Not mbln_BackupDatabase(clsOptions.NumDatabaseBackups) Then
          MsgBox Frame.ReadTextFile(Frame.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 600, 183), vbExclamation
        End If
    End If
        
    '..create the main front end object
    '
    ' 08 dec 08 - rg
    '
    If Not mbln_CreateObject(oCDM, DEF_OBJECT_VERSION & ".MainFrontEnd", DEF_PROJECT_DLL_NAME) Then GoTo Controlled_Exit
    'If Not mbln_CreateObject(aDOOR, "ADMFE.MainFrontEnd", DEF_PROJECT_DLL_NAME) Then GoTo Controlled_Exit

    '..try to get the default working unit
    If Len(GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_NESTING, "Units", "")) = 0 Then
        
        Load frmUnits
        frmUnits.Show
        
    End If
    
    DoEvents
    
    ' TFS#66387 - Alphacam crash on exit after CDM & AcamAddins.dll have been used
    ' Creating splash object here and passing through to CDM.dll seems pointless and
    ' was causing crash on 32Bit systems
    'Call m_Splash(oSplash)
    
    m_AlphaDOORRunning = True
    
    '..run processing (pass the program level for validation)
    oCDM.RunProcessing App '  ',oSplash   TFS#66387
    'oCDM.RunProcessing App, oSplash    ' 19 mar 11 - rg, modified to suit new object
    DoEvents
    
Controlled_Exit:

Call DebugNote("m_Processing Exit")

    m_AlphaDOORRunning = False
    
    'App.Unauthorise
    
    App.Frame.CloseProgressBox
                            
    Set oSplash = Nothing
    Set oCDM = Nothing
    
    '..unload all the forms                                                         '..07.17.02 - rg
    Call g_UnLoadAllForms
    
    If clsOptions.BackupDatabase And clsOptions.BackupDatabaseExit Then
        If Not mbln_BackupDatabase(clsOptions.NumDatabaseBackups) Then
          MsgBox Frame.ReadTextFile(Frame.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 600, 183), vbExclamation
        End If
    End If
    
    Set clsOptions = Nothing
    
Exit Function
    
m_Processing_Error:

    'App.Unauthorise
    
    If Err.Number = -2147417848 Then
        
Debug.Assert (Err.Number = 0)
        
        WriteError Err, True, "m_Processing"
        Resume Next
        
    Else
    
        MsgBox Err.Description, vbExclamation, "m_Processing"
        If (Err.Number <> 0) Then WriteError Err, True, "m_Processing"
        Resume Controlled_Exit
    
    End If

End Function

Function m_About()
      
    '..load the about form
    Load frmAbout
    frmAbout.Show
    
End Function

Function m_Help()

    LoadHelp

End Function

Private Sub m_Splash(oSplash As Object)
        
On Error Resume Next
    
    ' 18 mar 11 - rg
    '
    Set oSplash = CreateObject(DEF_OBJECT_VERSION & ".SplashScreen")    ' 19 mar 11 - rg, modified to suit new object

    '..just show little box if dll error
    If (Err.Number <> 0) Then
        
        With App.Frame
            .ShowProgressBox .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 120, 1), .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 300, 18)
        End With
        
        Err.Clear
        
    End If
    
End Sub

Private Function mbln_OKtoRun() As Boolean

    Dim FSO                     As New Scripting.FileSystemObject
    Dim fsoFile                 As Scripting.File
    Dim Fr                      As Frame
    Dim sCTX                    As String
    Dim bCancel                 As Boolean

On Error GoTo mbln_OKtoRun_Error

    '..assume we're ok
    mbln_OKtoRun = True
        
    Set Fr = App.Frame
    
    With Fr
    
        sCTX = gstr_CheckDir(.PathOfThisAddin) & DEF_TEXT
'        sDB = GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_OPTIONS, "DatabasePath", LicomdatPath & "Licomdat\AlphaDoor Data\")
'        'gstr_CheckDir(.PathOfThisAddin) & DATABASE_NAME

    End With
    
    ' Ensure the VBA current directory is set correctly
    ChDir App.Frame.PathOfThisAddin
    
    If Not gbln_ConnectToDB Then
        MsgBox Fr.ReadTextFile(sCTX, 500, 111) & Space(3), vbInformation, DEF_PROJECT_NAME
        mbln_OKtoRun = False
        GoTo Controlled_Exit
    End If
    
    
    ' Connected to database - now get the path
    sCurrentDB = gdb_CDM.Properties("Data Source Name").Value
    
    
    ' Test for SQL Server
    If gdb_CDM.Properties("DBMS Name").Value <> "Microsoft SQL Server" Then
    
        '..ensure that database is NOT read only
        If FSO.FileExists(sCurrentDB) Then
        
            Set fsoFile = FSO.GetFile(sCurrentDB)
            fsoFile.Attributes = 0 ' Normal
            
        Else
            
            '..try to get new database
            MsgBox Fr.ReadTextFile(sCTX, 500, 111) & Space(3), vbInformation, DEF_PROJECT_NAME
    ''        Call m_SelectDB(sDB, bCancel)
            
            If bCancel Then mbln_OKtoRun = False: GoTo Controlled_Exit
            
        End If
    Else
      gbln_SQLServer = True
    End If
    
    '..validate it
    Call m_ValidDatabase(bCancel)

    '..if canceled then bail
    If bCancel Then mbln_OKtoRun = False: GoTo Controlled_Exit

'    If Not CBool(GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_REGISTRATION, DEF_REG_KEY_ISUPDATED, 0)) Then
'
'        '..try to update
'        Call m_Update
'
'        '..check to ensure update was run
'        mbln_OKtoRun = CBool(GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_REGISTRATION, DEF_REG_KEY_ISUPDATED, 0))
'
'    End If

Controlled_Exit:

    Set fsoFile = Nothing
    Set FSO = Nothing
    Set Fr = Nothing

Exit Function

mbln_OKtoRun_Error:

    MsgBox Err.Description, vbExclamation, "mbln_OKtoRun"
    mbln_OKtoRun = False
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_OKtoRun"
    Resume Controlled_Exit

End Function

'Public Sub m_Update()
'
'    Dim aDOOR                   As Object
'
'On Error GoTo mbln_Update_Error
'
'' >< CDM
'
'    '..create the main front end object
'    If Not mbln_CreateObject(aDOOR, "ADUPD.UpdateADOOR", DEF_PROJECT_DLL_UPDATE_NAME) Then
'        GoTo Controlled_Exit
'    End If
'
'    '..run options (pass the program level for validation)
'    aDOOR.RunAlphaDOORUpdate App, DEF_MACRO_VERSION
'
'Controlled_Exit:
'
'    '..remove the main fron end object
'    Set aDOOR = Nothing
'
'Exit Sub
'
'mbln_Update_Error:
'
'    If (Err.Number <> 0) Then WriteError Err, True, "m_Update"
'    MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
'    Resume Controlled_Exit
'
'End Sub

Private Sub m_ValidDatabase(bCancel As Boolean)

        Dim Fr                      As Frame
        Dim rV                      As ADODB.Recordset
        Dim sMsg                    As String
        Dim bValid                  As Boolean
        Dim iReturnVal              As AdoorUpdateDatabaseFlag
        Dim dblV                    As Double
        Dim lngRet                  As Long
    
On Error Resume Next
                  
        ' 16 feb 12 TFS#48963
        '   + ROUTINE has been heavily modified to include additional database version check
                  
        bCancel = False
        bValid = True
                     
        If Not gbln_ConnectToDB Then Exit Sub
                                
        ' first check the macro version
        Set rV = gdb_CDM.Execute("SELECT Version FROM AD_VERSION")
        
        If (rV Is Nothing) Then
                bValid = False
        Else
                If (PDbl(gvar_CheckNull(rV.Fields!Version)) <> PDbl(DEF_MACRO_VERSION)) Then
                        bValid = False
                End If
        End If
        
        Call g_CloseRS(rV)
        Set rV = Nothing
        
        If bValid Then
                
                Set rV = gdb_CDM.Execute("SELECT * FROM AD_VERSION")

                ' add DatabaseVersion field if not there, set initial value to "0.0"
                If Not mbln_DBFieldExists(rV, "DatabaseVersion") Then
                        
                        Set rV = gdb_CDM.Execute("ALTER TABLE AD_VERSION ADD DatabaseVersion NVARCHAR(50)", lngRet)
                        Set rV = gdb_CDM.Execute("UPDATE AD_VERSION SET DatabaseVersion='0.0'")
                        Set rV = gdb_CDM.Execute("SELECT DatabaseVersion FROM AD_VERSION")
                        
                        If Not (rV Is Nothing) Then
                                If (rV.BOF And rV.EOF) Then Set rV = Nothing
                        End If
                        
                End If
        
                ' check again
                If (rV Is Nothing) Then
                        bValid = False
                Else
                        dblV = PDbl(gvar_CheckNull(rV.Fields!DatabaseVersion))
                End If
        
        End If
        
        Set Fr = App.Frame
    
        If Not bValid Then
                bCancel = True
                sMsg = Fr.ReadTextFile(Fr.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 107)
                MsgBox sMsg, vbExclamation, DEF_PROJECT_NAME
                GoTo Controlled_Exit
        End If
        
        If (PDbl(DEF_DB_VERSION) > dblV) Then
        
                iReturnVal = mint_UpdateDB
            
                ' let's see what was returned
                If (iReturnVal <> adoorUPDATE_DB_SUCCESS) Then
                        bCancel = True
                        sMsg = Fr.ReadTextFile(Fr.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 123) & " <Op: " & iReturnVal & ">"
                        MsgBox sMsg, vbExclamation, DEF_PROJECT_NAME
                Else
                        
                        ' now update the version number
                        Call g_CloseRS(rV)
                        Set rV = gdb_CDM.Execute("UPDATE AD_VERSION SET DatabaseVersion=" & gs_NoComma(DEF_DB_VERSION))
                        
                End If
                
        End If
    
Controlled_Exit:
    
        Call g_CloseRS(rV)
        Set rV = Nothing

Exit Sub

ErrTrap:
    
        bCancel = True
        MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
        If (Err.Number <> 0) Then WriteError Err, False, "m_ValidDatabase"
        Resume Controlled_Exit

End Sub

''Private Sub m_ValidDatabase(bCancel As Boolean)
''
''On Error GoTo m_ValidDatabase_Error
''
''    Dim rDate                   As ADODB.Recordset
''    Dim sMsg                    As String
''    Dim bValid                  As Boolean
''    Dim iReturnVal              As AdoorUpdateDatabaseFlag              ' altered from iEngravingReturnVal 07/25/03 - rg
''
''On Error Resume Next
''
''    bCancel = False
''
''    '..connect
''    If Not gbln_ConnectToDB Then Exit Sub
''
''    '..check need for update
'''    If Not CBool(GetSetting(DEF_REG_APP_NAME, DEF_REG_SECTION_REGISTRATION, DEF_REG_KEY_ISUPDATED, 0)) Then
'''
'''        '..try to update
'''        Call m_Update
'''
'''    End If
''
''    bValid = True
''
''    '..look for valid database
''    Set rDate = gdb_CDM.Execute("SELECT Version FROM AD_VERSION")
''
''    '..ensure version number exists
''    If (rDate Is Nothing) Then
''
''        bValid = False
''
''    Else
''
''        '..version date found, now compare against current
''        If Trim$(rDate.Fields!Version) <> Trim$(DEF_MACRO_VERSION) Then
''            bValid = False
''        End If
''
''    End If
''
''    If Not bValid Then
''        DoEvents
''
''        With App.Frame
''
''            sMsg = .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 107) & Space(3) & vbCrLf
''            sMsg = sMsg & .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 115) & Space(3) & vbCrLf
''            sMsg = sMsg & LicomdatPath & "Licomdat\AlphaDoor Data" & Space(3) & vbCrLf
''            sMsg = sMsg & vbCrLf
''            sMsg = sMsg & .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 116) & Space(3)
''
''            MsgBox sMsg, vbExclamation, DEF_PROJECT_NAME
''
''            bCancel = True
''
''        End With
''
''    End If
''
''    iReturnVal = mint_UpdateDB
''
''    ' let's see what was returned
''    If (iReturnVal <> adoorUPDATE_DB_SUCCESS) Then
''        With App.Frame
''            sMsg = .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 123) & " <Op: " & iReturnVal & ">"
''            MsgBox sMsg, vbExclamation, DEF_PROJECT_NAME
''        End With
''    End If
''
''' REPLACED WITH ABOVE 07/25/03 - rg
'''    If iEngravingReturnVal <> 0 Then
'''        sMSG = Frame.ReadTextFile(Frame.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 119) & "Op:" & iEngravingReturnVal
'''        MsgBox sMSG, vbExclamation, DEF_PROJECT_NAME
'''        bCancel = True
'''    End If
''
'''    '..if not up to date, then set flag to launch update util
'''    If Not bValid Then
'''
'''        '..not right
'''        DoEvents
'''
'''        With App.Frame
'''
'''            sMSG = .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 107) & Space(3) & vbCrLf
'''            sMSG = sMSG & .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 108) & Space(3)
'''
'''        End With
'''
'''        If MsgBox(sMSG, vbQuestion + vbYesNo, DEF_PROJECT_NAME) = vbYes Then
'''            SaveSetting DEF_REG_APP_NAME, DEF_REG_SECTION_REGISTRATION, DEF_REG_KEY_ISUPDATED, 0
'''            bCancel = False
'''        Else
'''            bCancel = True
'''        End If
'''
'''    End If
''
''Controlled_Exit:
''
''    Call g_CloseRS(rDate)                                       ' 07/25/03 - rg
''    Set rDate = Nothing
''
''Exit Sub
''
''m_ValidDatabase_Error:
''
''    MsgBox Err.Description, vbExclamation, DEF_PROJECT_NAME
''    If (Err.Number <> 0) Then WriteError Err, False, "m_ValidDatabase"
''    Resume Controlled_Exit
''
''End Sub

Private Function mbln_CreateObject(oDOOR As Object, sObject As String, sDLL As String) As Boolean
 
On Error Resume Next
    
    mbln_CreateObject = True
    
    '..try to create the front end object
    Set oDOOR = CreateObject(sObject)
    
    '..look for error
    If Err.Number > 0 Then
        
        '..found an error so check the number
        If Err.Number = 429 Then
                                        
            If Not RegSvr32(App.Frame.PathOfThisAddin & DEF_BACKSLASH & sDLL, False) Then
                
                DoEvents
                
                '..bummer
                mbln_CreateObject = False
                GoTo mbln_CreateObject_Error
                
            Else
                
                '..should not be registered so create the object
                Set oDOOR = CreateObject(sObject)
                
            End If
        
        Else
            
            '..something else must be wrong
            GoTo mbln_CreateObject_Error
            
        End If
        
    End If
    
Controlled_Exit:

Exit Function

mbln_CreateObject_Error:
    
    '..let the user know there is a serious problem
    With App.Frame
    
        MsgBox .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 81) & Space(3) & vbCrLf & _
               .ReadTextFile(.PathOfThisAddin & DEF_BACKSLASH & DEF_TEXT, 500, 82) & Space(3), vbCritical, DEF_PROJECT_NAME
           
    End With
           
    If (Err.Number <> 0) Then WriteError Err, True, "mbln_CreateObject"
            
End Function

Private Function mbln_ValidReference() As Boolean
    
    ' ************************************************************************************************
    ' *        Name: gf_ValidReference                                                               *
    ' *    Function: Validates all references used by this project (Macro)                           *
    ' *              returns TRUE if all references are valid, FALSE otherwise.                      *
    ' *  Written By: Mike George mg@us.licom.com                                                     *
    ' *        Date: 07-13-00                                                                        *
    ' *  ' xx Modified for AlphaCABINET 4/30/2000                                                    *
    ' ************************************************************************************************
    
    Dim VBPs                    As VBProjects
    Dim VBP                     As VBProject
    Dim obj_References          As References
    Dim obj_Reference           As Reference

On Error GoTo mbln_ValidReference_Err
    
    ' -- Assume success
    mbln_ValidReference = True
    
    ' -- get all projects
    Set VBPs = Application.VBE.VBProjects
    
    ' -- iterate each project
    For Each VBP In VBPs
        
        ' -- Get the Project we want
        If StrComp(VBP.Name, DEF_PROJECT_NAME, vbTextCompare) = 0 Then
            
            ' -- get all references
            Set obj_References = VBP.References
            
            ' -- check that they exist and are valid
            For Each obj_Reference In obj_References
                
                ' -- We have a problem ?
                If obj_Reference.IsBroken Then
                    ' -- Er....
                    MsgBox "Reference " & obj_Reference.Name & " is missing, unable to load macro.", vbCritical, DEF_PROJECT_NAME
                    mbln_ValidReference = False
                    Exit Function
                End If
                
            Next
            
        End If
        
    Next
        
ControlledExit:
    
    ' -- clean up
    Set VBPs = Nothing
    Set VBP = Nothing
    Set obj_References = Nothing
    Set obj_Reference = Nothing
       
Exit Function
    
mbln_ValidReference_Err:
    
    MsgBox "An error has occured validating references for " & DEF_PROJECT_NAME & ".", vbExclamation, DEF_PROJECT_NAME
    mbln_ValidReference = False
    Resume ControlledExit
    
End Function

Private Function RegSvr32(ByVal FileName As String, bUnReg As Boolean) As Boolean

    Dim lLib                As Long
    Dim lProcAddress        As Long
    Dim lSuccess            As Long
    Dim lExitCode           As Long
    Dim lThread             As Long
    Dim bAns                As Boolean
    Dim sPurpose            As String
    Dim FSO                 As New Scripting.FileSystemObject
    
    sPurpose = IIf(bUnReg, "DllUnregisterServer", "DllRegisterServer")
    
    If Not FSO.FileExists(FileName) Then GoTo Controlled_Exit
    
    lLib = LoadLibraryRegister(FileName)
    
    'could load file
    If lLib = 0 Then GoTo Controlled_Exit
    
    lProcAddress = GetProcAddressRegister(lLib, sPurpose)
    
    If lProcAddress = 0 Then
      
      'Not an ActiveX Component
       FreeLibraryRegister lLib
       GoTo Controlled_Exit
       
    Else
       
       lThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lThread)
       
       If lThread Then
            
            lSuccess = (WaitForSingleObject(lThread, 10000) = 0)
            
            If Not CBool(lSuccess) Then
            
               Call GetExitCodeThread(lThread, lExitCode)
               Call ExitThread(lExitCode)
               bAns = False
               
               GoTo Controlled_Exit
            
            Else
               bAns = True
            End If
            
            CloseHandle lThread
            FreeLibraryRegister lLib
       
       End If
       
    End If
        
    RegSvr32 = bAns
    
Controlled_Exit:
    
    Set FSO = Nothing
    
End Function

' 15 dec 08 - rg
'
Public Sub g_ConfigDB()
        
        Dim FSO                     As New Scripting.FileSystemObject
        Dim lngRet                  As Long
        Dim strUDL                  As String
        Dim strCTX                  As String
        
        Const SW_SHOWNORMAL         As Long = 1
        
On Error Resume Next
        
        ' build path to UDL file
        'strUDL = gs_ThisDir & "CDM.udl"
        
        g_MoveUDL strUDL
        
        If FSO.FileExists(strUDL) Then
                lngRet = ShellExecute(App.Frame.WindowHandle, "open", strUDL, vbNullString, vbNullString, SW_SHOWNORMAL)
        Else
                
                strCTX = gs_ThisDir & DEF_TEXT
                
                ' Cannot find datalink file.
                MsgBox App.Frame.ReadTextFile(strCTX, 500, 2)
                
        End If
        
        Set FSO = Nothing

End Sub

