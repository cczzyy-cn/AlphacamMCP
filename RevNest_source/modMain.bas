in file: C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\000\StartUp\Utils\ReverseNest\ReverseNest.amb - OLE stream: 'vao/The VBA Project/_VBA_Project/VBA/modMain'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

Private Const ATT_PATH_FILE             As String = "LicomUKsab_nest_path_file"
Private Const ATT_FIRST_PATH            As String = "LicomUKsab_nest_first_path"
Private Const ATT_REQUIRED              As String = "LicomUKsab_nest_required"
Private Const ATT_SHEET_IDENT           As String = "LicomUKsab_sheet_ident"
Private Const ATT_SHEET_MATERIAL        As String = "LicomUKsab_sheet_material"
Private Const ATT_SHEET_THICKNESS       As String = "LicomUKsab_sheet_thickness"
Private Const ATT_PART_MOVEX            As String = "LicomUKsab_part_movex"
Private Const ATT_PART_MOVEY            As String = "LicomUKsab_part_movey"
Private Const ATT_PART_ROTANGLE         As String = "LicomUKsab_part_rotangle"
Private Const ATT_PART_MIRRORED         As String = "LicomUKsab_part_mirrored"
Private Const ATT_IS_BOBBLE             As String = "LicomUKsab_is_bobble"

Private Const ATT_PART_MOVE_BY_X        As String = "LicomUKja_part_move_by_x"
Private Const ATT_PART_MOVE_BY_Y        As String = "LicomUKja_part_move_by_y"
Private Const ATT_PART_SHIFT_X          As String = "LicomUKja_part_shift_x"
Private Const ATT_PART_SHIFT_Y          As String = "LicomUKja_part_shift_y"

'01 OCT 10  +SDO
Private Const ATT_NEST_ITEM_NUM         As String = "LicomUKsab_nest_item_number"
Private Const ATT_IS_REV_SIDE           As String = "AcamUSrg_IsReverseSide"
Private Const ATT_REV_TEXT              As String = "AcamUSrg_TextIsReversed"
'

Public Sub g_AroundX(ByVal bSheetOrder As Boolean, ByVal bMinToolChanges As Boolean)

On Error Resume Next
        
    ' Variables
    Dim Drw As Drawing
    Dim tmpdrw As Drawing
    Dim P As Path
    Dim pcopy As Path
    Dim newp As Path
    Dim lastsheet As Path
    Dim prefix As String
    Dim suffix As String
    Dim strName As String
    Dim minx As Double
    Dim maxx As Double
    Dim miny As Double
    Dim maxy As Double
    Dim mirrorx As Double
    Dim xmove As Double
    Dim ymove As Double
    Dim rotate As Double
    Dim count As Integer
    Dim count2 As Integer
    Dim reflect As Integer
    Dim grp As Integer
    Dim isfirst As Long
    Dim elem As Element
    Dim high As Double
    Dim big As Double
    Dim exmin As Double
    Dim exmax As Double
    Dim eymin As Double
    Dim eymax As Double
    Dim textx As Double
    Dim texty As Double
    Dim tp As Path
    Dim coll As Paths
    Dim wrd As String
    Dim ni As NestInformation
    Dim sh As NestSheet
    Dim inst As NestPartInstance
    Dim I As Integer
    Dim lastop As Integer
    Dim sheetop As Integer
    Dim maxop As Integer
    Dim minop As Integer
    Dim currtool As MillTool
    Dim strReverse As String
    Dim T As Text
    Dim T2 As Text
    Dim pT As Path
    Dim psT As Paths
    Dim dblX As Double
    
    Set Drw = App.ActiveDrawing
    Set ni = Drw.GetNestInformation
    
    ' Some basic tests...
    If ni.Sheets.count = 0 Then
        MsgBox gv_CTX(30, 1, "No nesting found in current drawing"), vbInformation
        GoTo byebye
    End If
    
    If Drw.GetToolPathCount = 0 Then
        MsgBox gv_CTX(30, 2, "No front-side toolpaths found."), vbInformation
        GoTo byebye
    End If
    
    Call g_LockAcam
    
    ' Start by finding the extents of the sheets, so that
    ' we can know where our mirror line lies
    minx = miny = 1E+20
    maxx = maxy = -1E+20
    
    For Each sh In ni.Sheets
        Set P = sh.Geometry
        If (P.MinXL < minx) Then minx = P.MinXL
        If (P.MinYL < miny) Then miny = P.MinYL
        If (P.MaxXL > maxx) Then maxx = P.MaxXL
        If (P.MaxYL > maxy) Then maxy = P.MaxYL
    Next sh
        
    strReverse = gv_CTX(30, 3, "(reverse)")
        
' mirror around the far right
    
    ' Okay, we now know the extents of the sheets, so mirror
    ' them 5% below the bottom of the lowermost.
    mirrorx = minx - ((maxx - minx) * 0.05)
    
    ' Now set about copying and mirroring the sheets.
    ' We also mirror the identity bubbles, but NOT the
    ' text inside.
    Set lastsheet = Nothing
    Set P = Drw.GetFirstGeo
    For count = Drw.GetGeoCount To 1 Step -1
        If (P.Sheet Or P.Dimension) Then
            ' If it's dimension, then make sure it's
            ' the line or the circle, not the text.
            If (P.Dimension) Then
                ' If it has no "it's the bobble" attribute, then
                ' we musn't mirror it.
                If (P.Attribute(ATT_IS_BOBBLE) = 0) Then GoTo loopnext
            End If
            
            ' Copy the path and mirror it
            Set pcopy = P.CopyTemporary
            pcopy.MirrorL mirrorx, miny, mirrorx, maxy
            pcopy.StoreTemporary
            
            ' Change the attribute to specify that it's the
            ' reverse-side sheet.
            If (P.Sheet) Then
                Set lastsheet = P
                grp = Drw.GetNextGroupNumberForGeometries
                
                strName = pcopy.Attribute(ATT_SHEET_IDENT)

                ' 01 oct 10 - rg
                '   + REPLACED "(reverse)" suffix with hard-coded " rev"
                '
                'strName = strName + " " + strReverse
                strName = strName & " rev"
                
                ' Copy out all the attributes which we'll need.
                pcopy.Attribute(ATT_SHEET_IDENT) = strName
                pcopy.Attribute(ATT_SHEET_MATERIAL) = P.Attribute(ATT_SHEET_MATERIAL)
                pcopy.Attribute(ATT_SHEET_THICKNESS) = P.Attribute(ATT_SHEET_THICKNESS)
                
                ' 01 oct 10 - rg
                '
                pcopy.Attribute(ATT_IS_REV_SIDE) = 1
                '
                ' create item numbers
                For Each T In App.ActiveDrawing.Text
                        
                        If (T.Attribute(ATT_REV_TEXT) = 0) Then
                        
                                Set psT = T.ConvertToTemporaryGeometry
                                
                                If (psT(1).TestInsidePath(P) = acamResultTRUE) Then
                                        
                                        For Each pT In psT
                                                Call pT.MirrorL(mirrorx, miny, mirrorx, maxy)
                                        Next pT
                                        
                                        Call psT.GetExtentL(dblX, 0, 0, 0)
                                        
                                        Set T2 = T.Copy
                                        
                                        Call T2.MoveL((dblX - T.MinXL), 0)
                                        
                                        ' set flags so we don't bother again
                                        T.Attribute(ATT_REV_TEXT) = 1
                                        T2.Attribute(ATT_REV_TEXT) = 1
                                        T2.Attribute(ATT_IS_REV_SIDE) = 1
                                        
                                        ' update sheet ID
                                        T2.Attribute(ATT_SHEET_IDENT) = strName
                                
                                End If
                                
                                Call psT.Delete
                                
                        End If
                
                Next T
                
            Else
                ' If it's the circle, then that's what we
                ' put text in.
                Set elem = P.GetFirstElem
                If elem.IsArc Then
                    ' It's a circle, so write the relevant text
                    ' inside it. The text is the sheet's name, but
                    ' with lower case.
                    wrd = lastsheet.Attribute(ATT_SHEET_IDENT)
                    wrd = LCase(wrd)
                    
                    high = pcopy.MaxYL - pcopy.MinYL
                    
                    ' Create some text, so we can determine sizings.
                    Set coll = Drw.CreateText(wrd, 0, 0, high)
                    
                    ' Find the resizing factor, then delete the collection
                    ' This involves finding the extents of the text.
                    exmin = eymin = 1E+20
                    exmax = eymax = -1E+20
                    For Each tp In coll
                        If tp.MinXL < exmin Then exmin = tp.MinXL
                        If tp.MinYL < eymin Then eymin = tp.MinYL
                        If tp.MaxXL > exmax Then exmax = tp.MaxXL
                        If tp.MaxYL > eymax Then eymax = tp.MaxYL
                    Next
                    
                    ' 'big' becomes the largest dimension
                    If (exmax - exmin) > (eymax - eymin) Then
                        big = exmax - exmin
                    Else
                        big = eymax - eymin
                    End If
                    
                    coll.Selected = True
                    Drw.DeleteSelected
                    
                    Dim SF As Double
                    SF = (0.707 * high / big)
                    
                    textx = pcopy.MinXL + ((pcopy.MaxXL - pcopy.MinXL) / 2)
                    texty = pcopy.MinYL + ((pcopy.MaxYL - pcopy.MinYL) / 2)
                    textx = textx - (((exmax - exmin) / 2) * SF)
                    texty = texty - (((eymax - eymin) / 2) * SF)
                    
                    ' Okay, then rescale the height and draw the real text.
                    high = high * SF
                    Set coll = Drw.CreateText(wrd, textx, texty, high)
                    
                    For Each tp In coll
                        tp.Group = grp
                        tp.Dimension = True
                    Next
                End If
            End If
            
            ' And make sure the group number changes as well.
            pcopy.Group = grp
        End If
        
loopnext:
        Set P = P.GetNext   ' On to the next path!
    Next
    'Drw.ZoomAll

    ' Now the harder job - mirroring the toolpaths/geometries.
    ' Run through sheet by sheet.
    lastop = Drw.Operations.count + 1
    For Each sh In ni.Sheets
        ' If we need to order by sheet, then do it now so that inserted
        ' paths fall in place automatically.
        If bSheetOrder Then
            ' Find the largest operation number and the smallest in the sheet.
            maxop = 0
            minop = lastop
            For Each tp In sh.Paths
                If maxop < tp.OpNo Then maxop = tp.OpNo
                If minop > tp.OpNo Then minop = tp.OpNo
            Next tp
            
            ' Now move all those operations to the end.
            For I = minop To maxop
                Drw.Operations.Renumber minop, lastop, acamOpADD_TO_OPERATION
            Next I
        End If

        ' Now insert the reverse-side parts for this sheet.
        sheetop = lastop      ' The insert point for this sheet.
        
        ' Now we can run through the paths on this sheet and
        ' mirror each of them safely.
        ' We only care to do this with the "first path"
        ' paths, since the rest come along with the file.
        For Each inst In sh.Parts
            Set P = inst.Paths(1)
            If P.Attribute(ATT_FIRST_PATH) = 0 Then GoTo loopagain  ' Sanity check
            
            ' Query the path's filename
            strName = P.Attribute(ATT_PATH_FILE)
            
            ' Add the reverse-side suffix. Remember that it
            ' must go BEFORE the ".amd" or whatever the prefix is!
            prefix = Left(strName, Len(strName) - 4)
            suffix = Right(strName, 4)
            strName = prefix + "_rev" + suffix
            
            ' Now see if we can load up the new name!
            Set tmpdrw = Nothing
            Set tmpdrw = App.OpenTempDrawing(strName)
            
            If Not tmpdrw Is Nothing Then
                    
                ' TFS#80910 - test to ensure the drawing we are inserting does not contain workplanes
                If tmpdrw.WorkPlanes.count > 0 Then
                  MsgBox strName & Chr(13) & gv_CTX(30, 5, "contains workplanes and cannot be used"), vbExclamation
                  GoTo loopagain
                End If
                    
                ' Excellent - we got it. Now the slightly more tricky
                ' part... We extract the movement/rotation information for the
                ' part, rotate, move the new part to there then mirror it.
                Set newp = tmpdrw.GetFirstToolPath
                isfirst = 1
          
                ' Then we can go through the paths and do the rotation
                ' and translation etc.
                For count2 = 1 To tmpdrw.GetToolPathCount
                    
                    '01 OCT 10 +-SDO Move most of the path manipulation into it own Function
                    'Function SetAttribs to make it accessable when wanting to mirror normal
                    'geometries
                    Call SetAttribs(P, newp, isfirst, mirrorx, miny, maxy, reflect, True, strName)

                    If isfirst = 1 Then isfirst = 0
                    
                    ' Copy it and store in the main drawing. We re-set the
                    ' operation to either the last one in the list, or one
                    ' which shares our tool, depending on the option.
                    Set pcopy = newp.CopyTemporary
                    
                    If Not bMinToolChanges Then
                        pcopy.OpNo = lastop
                        lastop = lastop + 1
                    Else
                        ' Find a tool with the same number in the recently
                        ' added paths.
                        Set currtool = pcopy.GetTool
                        For I = sheetop To lastop - 1
                            If currtool.Number = Drw.Operations(I).Tool.Number Then Exit For
                        Next I
                        pcopy.OpNo = I
                        
                        If I = lastop Then lastop = lastop + 1
                    End If
                    
                    pcopy.StoreTemporary
                    
                    ' On to the next path!
                    Set newp = newp.GetNext
                Next count2
                
                '01 OCT 10 +SDO Added the option to include normal Geos when generating the reverse side sheet
                If frmMain.chkGeos.Value = True Then
                        For Each newp In tmpdrw.Geometries
                        
                                If Not newp.IsToolPath And Not newp.Sheet Then
                                        Call SetAttribs(P, newp, isfirst, mirrorx, miny, maxy, reflect, True, strName)
                                        Set pcopy = newp.CopyTemporary
                                        pcopy.StoreTemporary
                                End If

                        Next
                End If
  
            End If
loopagain:
        Next inst
loopalso:
    Next sh
    Drw.Operations.OrderAll     ' Re-order operations correctly.
    'Drw.Redraw
   
byebye:
        
    Set Drw = Nothing
    Set tmpdrw = Nothing
    Set P = Nothing
    Set pcopy = Nothing
    Set newp = Nothing
    Set lastsheet = Nothing
    Set elem = Nothing
    Set tp = Nothing
    Set coll = Nothing
    Set ni = Nothing
    Set sh = Nothing
    Set inst = Nothing
    Set currtool = Nothing
    Set T = Nothing
    Set pT = Nothing
    Set psT = Nothing
    
    Call g_UnlockAcam(True)
    
End Sub

Public Sub g_AroundY(ByVal bSheetOrder As Boolean, ByVal bMinToolChanges As Boolean)

On Error Resume Next

    ' Variables
    Dim Drw As Drawing
    Dim tmpdrw As Drawing
    Dim P As Path
    Dim pcopy As Path
    Dim newp As Path
    Dim lastsheet As Path
    Dim prefix As String
    Dim suffix As String
    Dim strName As String
    Dim minx As Double
    Dim maxx As Double
    Dim miny As Double
    Dim maxy As Double
    Dim mirrory As Double
    Dim xmove As Double
    Dim ymove As Double
    Dim rotate As Double
    Dim count As Integer
    Dim count2 As Integer
    Dim reflect As Integer
    Dim grp As Integer
    Dim isfirst As Long
    Dim elem As Element
    Dim high As Double
    Dim big As Double
    Dim exmin As Double
    Dim exmax As Double
    Dim eymin As Double
    Dim eymax As Double
    Dim textx As Double
    Dim texty As Double
    Dim tp As Path
    Dim coll As Paths
    Dim wrd As String
    Dim ni As NestInformation
    Dim sh As NestSheet
    Dim inst As NestPartInstance
    Dim I As Integer
    Dim lastop As Integer
    Dim sheetop As Integer
    Dim maxop As Integer
    Dim minop As Integer
    Dim currtool As MillTool
    Dim strReverse As String
    Dim T As Text
    Dim T2 As Text
    Dim pT As Path
    Dim psT As Paths
    Dim dblY As Double

    Set Drw = App.ActiveDrawing
    Set ni = Drw.GetNestInformation

    ' Some basic tests...
    If ni.Sheets.count = 0 Then
        MsgBox gv_CTX(30, 1, "No nesting found in current drawing"), vbInformation
        GoTo byebye
    End If

    If Drw.GetToolPathCount = 0 Then
        MsgBox gv_CTX(30, 2, "No front-side toolpaths found."), vbInformation
        GoTo byebye
    End If

    Call g_LockAcam

    ' Start by finding the extents of the sheets, so that
    ' we can know where our mirror line lies
    minx = miny = 1E+20
    maxx = maxy = -1E+20

    For Each sh In ni.Sheets
        Set P = sh.Geometry
        If (P.MinXL < minx) Then minx = P.MinXL
        If (P.MinYL < miny) Then miny = P.MinYL
        If (P.MaxXL > maxx) Then maxx = P.MaxXL
        If (P.MaxYL > maxy) Then maxy = P.MaxYL
    Next sh

    strReverse = gv_CTX(30, 3, "(reverse)")

    ' Okay, we now know the extents of the sheets, so mirror
    ' them 5% below the bottom of the lowermost.
    mirrory = miny - ((maxy - miny) * 0.05)

    ' Now set about copying and mirroring the sheets.
    ' We also mirror the identity bubbles, but NOT the
    ' text inside.
    Set lastsheet = Nothing
    Set P = Drw.GetFirstGeo
    For count = Drw.GetGeoCount To 1 Step -1
        If (P.Sheet Or P.Dimension) Then
            ' If it's dimension, then make sure it's
            ' the line or the circle, not the text.
            If (P.Dimension) Then
                ' If it has no "it's the bobble" attribute, then
                ' we musn't mirror it.
                If (P.Attribute(ATT_IS_BOBBLE) = 0) Then GoTo loopnext
            End If

            ' Copy the path and mirror it
            Set pcopy = P.CopyTemporary
            pcopy.MirrorL minx, mirrory, maxx, mirrory
            pcopy.StoreTemporary

            ' Change the attribute to specify that it's the
            ' reverse-side sheet.
            If (P.Sheet) Then
            
                Set lastsheet = P
                grp = Drw.GetNextGroupNumberForGeometries
                strName = pcopy.Attribute(ATT_SHEET_IDENT)
                
                ' 01 oct 10 - rg
                '   + REMOVED "(reverse)" suffix and enforced lower case sheet name
                '
                'strName = strName + " " + strReverse
''                strName = "Rev " & LCase$(strName)
                strName = strName & " rev"
                'strName = LCase$(strName)   ' + " " + strReverse

                ' Copy out all the attributes which we'll need.
                pcopy.Attribute(ATT_SHEET_IDENT) = strName
                pcopy.Attribute(ATT_SHEET_MATERIAL) = P.Attribute(ATT_SHEET_MATERIAL)
                pcopy.Attribute(ATT_SHEET_THICKNESS) = P.Attribute(ATT_SHEET_THICKNESS)
                
                ' 01 oct 10 - rg
                '
                pcopy.Attribute(ATT_IS_REV_SIDE) = 1
                '
                ' create item numbers
                For Each T In App.ActiveDrawing.Text
                        
                        If (T.Attribute(ATT_REV_TEXT) = 0) Then
                        
                                Set psT = T.ConvertToTemporaryGeometry
                                
                                If (psT(1).TestInsidePath(P) = acamResultTRUE) Then
                                        
                                        For Each pT In psT
                                                Call pT.MirrorL(minx, mirrory, maxx, mirrory)
                                        Next pT
                                        
                                        Call psT.GetExtentL(0, dblY, 0, 0)
                                        
                                        Set T2 = T.Copy
                                        
                                        Call T2.MoveL(0, (dblY - T.MinYL))
                                        
                                        ' set flags so we don't bother again
                                        T.Attribute(ATT_REV_TEXT) = 1
                                        T2.Attribute(ATT_REV_TEXT) = 1
                                        T2.Attribute(ATT_IS_REV_SIDE) = 1
                                        
                                        ' update sheet ID
                                        T2.Attribute(ATT_SHEET_IDENT) = strName
                                
                                End If
                                
                                Call psT.Delete
                                
                        End If
                
                Next T

            Else
                ' If it's the circle, then that's what we
                ' put text in.
                Set elem = P.GetFirstElem
                If elem.IsArc Then
                    ' It's a circle, so write the relevant text
                    ' inside it. The text is the sheet's name, but
                    ' with lower case.
                    wrd = lastsheet.Attribute(ATT_SHEET_IDENT)
                    wrd = LCase(wrd)

                    high = pcopy.MaxYL - pcopy.MinYL

                    ' Create some text, so we can determine sizings.
                    Set coll = Drw.CreateText(wrd, 0, 0, high)

                    ' Find the resizing factor, then delete the collection
                    ' This involves finding the extents of the text.
                    exmin = eymin = 1E+20
                    exmax = eymax = -1E+20
                    For Each tp In coll
                        If tp.MinXL < exmin Then exmin = tp.MinXL
                        If tp.MinYL < eymin Then eymin = tp.MinYL
                        If tp.MaxXL > exmax Then exmax = tp.MaxXL
                        If tp.MaxYL > eymax Then eymax = tp.MaxYL
                    Next

                    ' 'big' becomes the largest dimension
                    If (exmax - exmin) > (eymax - eymin) Then
                        big = exmax - exmin
                    Else
                        big = eymax - eymin
                    End If

                    coll.Selected = True
                    Drw.DeleteSelected

                    Dim SF As Double
                    SF = (0.707 * high / big)

                    textx = pcopy.MinXL + ((pcopy.MaxXL - pcopy.MinXL) / 2)
                    texty = pcopy.MinYL + ((pcopy.MaxYL - pcopy.MinYL) / 2)
                    textx = textx - (((exmax - exmin) / 2) * SF)
                    texty = texty - (((eymax - eymin) / 2) * SF)

                    ' Okay, then rescale the height and draw the real text.
                    high = high * SF
                    Set coll = Drw.CreateText(wrd, textx, texty, high)

                    For Each tp In coll
                        tp.Group = grp
                        tp.Dimension = True
                    Next
                End If
            End If

            ' And make sure the group number changes as well.
            pcopy.Group = grp
        End If

loopnext:
        Set P = P.GetNext   ' On to the next path!
    Next
    'Drw.ZoomAll

    ' Now the harder job - mirroring the toolpaths/geometries.
    ' Run through sheet by sheet.
    lastop = Drw.Operations.count + 1
    For Each sh In ni.Sheets
        ' If we need to order by sheet, then do it now so that inserted
        ' paths fall in place automatically.
        If bSheetOrder Then
            ' Find the largest operation number and the smallest in the sheet.
            maxop = 0
            minop = lastop
            For Each tp In sh.Paths
                If maxop < tp.OpNo Then maxop = tp.OpNo
                If minop > tp.OpNo Then minop = tp.OpNo
            Next tp

            ' Now move all those operations to the end.
            For I = minop To maxop
                Drw.Operations.Renumber minop, lastop, acamOpADD_TO_OPERATION   'Here
            Next I
        End If

        ' Now insert the reverse-side parts for this sheet.
        sheetop = lastop      ' The insert point for this sheet.

        ' Now we can run through the paths on this sheet and
        ' mirror each of them safely.
        ' We only care to do this with the "first path"
        ' paths, since the rest come along with the file.
        For Each inst In sh.Parts
            Set P = inst.Paths(1)
            If P.Attribute(ATT_FIRST_PATH) = 0 Then GoTo loopagain  ' Sanity check  'Here
            'Here
            ' Query the path's filename
            strName = P.Attribute(ATT_PATH_FILE)

            ' Add the reverse-side suffix. Remember that it
            ' must go BEFORE the ".amd" or whatever the prefix is!
            prefix = Left(strName, Len(strName) - 4)
            suffix = Right(strName, 4)
            strName = prefix + "_rev" + suffix

            ' Now see if we can load up the new name!
            Set tmpdrw = Nothing
            Set tmpdrw = App.OpenTempDrawing(strName)

            If Not tmpdrw Is Nothing Then

                ' Excellent - we got it. Now the slightly more tricky
                ' part... We extract the movement/rotation information for the
                ' part, rotate, move the new part to there then mirror it.
                
                ' TFS#80910 - test to ensure the drawing we are inserting does not contain workplanes
                If tmpdrw.WorkPlanes.count > 0 Then
                  MsgBox strName & Chr(13) & gv_CTX(30, 5, "contains workplanes and cannot be used"), vbExclamation
                  GoTo loopagain
                End If

                Set newp = tmpdrw.GetFirstToolPath
                isfirst = 1
 
                ' Then we can go through the paths and do the rotation
                ' and translation etc.
                For count2 = 1 To tmpdrw.GetToolPathCount

                    '01 OCT 10 +-SDO Move most of the path manipulation into it own Function
                    'Function SetAttribs to make it accessable when wanting to mirror normal
                    'geometries
                    Call SetAttribs(P, newp, isfirst, mirrory, minx, maxx, reflect, False, strName)
                    
                    If isfirst = 1 Then isfirst = 0

                    ' Copy it and store in the main drawing. We re-set the
                    ' operation to either the last one in the list, or one
                    ' which shares our tool, depending on the option.
                    Set pcopy = newp.CopyTemporary

                    If Not bMinToolChanges Then
                        pcopy.OpNo = lastop
                        lastop = lastop + 1
                    Else
                        ' Find a tool with the same number in the recently
                        ' added paths.
                        Set currtool = pcopy.GetTool
                        For I = sheetop To lastop - 1
                            If currtool.Number = Drw.Operations(I).Tool.Number Then Exit For
                        Next I
                        pcopy.OpNo = I

                        If I = lastop Then lastop = lastop + 1
                    End If

                    pcopy.StoreTemporary

                    ' On to the next path!
                    Set newp = newp.GetNext
                Next count2
                
                '01 OCT 10 +SDO Added the option to include normal Geos when generating the reverse side sheet
                If frmMain.chkGeos.Value = True Then
                        For Each newp In tmpdrw.Geometries
                        
                                If Not newp.IsToolPath And Not newp.Sheet Then
                                        Call SetAttribs(P, newp, isfirst, mirrory, minx, maxx, reflect, False, strName)
                                        Set pcopy = newp.CopyTemporary
                                        pcopy.StoreTemporary
                                End If

                        Next
                End If
                
                
            End If
loopagain:
        Next inst
loopalso:
    Next sh
    Drw.Operations.OrderAll     ' Re-order operations correctly.
    'Drw.Redraw

byebye:

    Set Drw = Nothing
    Set tmpdrw = Nothing
    Set P = Nothing
    Set pcopy = Nothing
    Set newp = Nothing
    Set lastsheet = Nothing
    Set elem = Nothing
    Set tp = Nothing
    Set coll = Nothing
    Set ni = Nothing
    Set sh = Nothing
    Set inst = Nothing
    Set currtool = Nothing
    Set T = Nothing
    Set pT = Nothing
    Set psT = Nothing

    Call g_UnlockAcam(True)

End Sub

'01 OCT 10 +SDO Moved most of the path manipulation to this function to encapsulate the code
'so as to use it across all path types.
Private Function SetAttribs(orgPath As Path, curPth As Path, isFirstTP As Long, dblMirror As Double, _
                            dblMin As Double, dblMax As Double, intReflect As Integer, blnAboutX As Boolean, sName As String) As Boolean
    
    Dim xmove As Double
    Dim ymove As Double
    Dim rotate As Double
    Dim reflect As Integer
    
                rotate = orgPath.Attribute(ATT_PART_ROTANGLE)
                xmove = orgPath.Attribute(ATT_PART_MOVEX)
                ymove = orgPath.Attribute(ATT_PART_MOVEY)
                reflect = orgPath.Attribute(ATT_PART_MIRRORED)


                Dim MoveX As Double
                Dim MoveY As Double
                Dim ShiftX As Double
                Dim ShiftY As Double

                MoveX = orgPath.Attribute(ATT_PART_MOVE_BY_X)
                MoveY = orgPath.Attribute(ATT_PART_MOVE_BY_Y)
                ShiftX = orgPath.Attribute(ATT_PART_SHIFT_X)
                ShiftY = orgPath.Attribute(ATT_PART_SHIFT_Y)
                
                curPth.MoveL ShiftX, ShiftY

                ' We reflect BEFORE rotating!
                If intReflect = 1 Then
                    curPth.MirrorL 0, 1, 0, 0
                    rotate = -rotate    ' Reflection means opposite direction rotation!
                End If
            
                ' Now rotate
                curPth.RotateL rotate, 0, 0

                ' Then move and reflect into place.
                curPth.MoveL MoveX, MoveY
                
                If blnAboutX Then
                        curPth.MirrorL dblMirror, dblMin, dblMirror, dblMax
                Else
                        curPth.MirrorL dblMin, dblMirror, dblMax, dblMirror
                End If
                
                'newp.Redraw

                ' Set all the appropriate attributes.
                curPth.Attribute(ATT_FIRST_PATH) = isFirstTP
                curPth.Attribute(ATT_PATH_FILE) = sName
                curPth.Attribute(ATT_REQUIRED) = orgPath.Attribute(ATT_REQUIRED)
                
                '01 OCT 10 +SDO
                curPth.Attribute(ATT_NEST_ITEM_NUM) = orgPath.Attribute(ATT_NEST_ITEM_NUM)
                
                ' 01 oct 10 - rg
                '
                curPth.Attribute(ATT_IS_REV_SIDE) = 1
                
End Function
-------------------------------------------------------------------------------
