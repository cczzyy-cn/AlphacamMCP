in file: C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\000\StartUp\Utils\ReverseNest\ReverseNest.amb - OLE stream: 'vao/The VBA Project/_VBA_Project/VBA/modGeneral'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

' >< ENUMS ><
'
Public Enum SystemRegionalSetting
        RegionalSetting_DATE_SEPARATOR
        RegionalSetting_DECIMAL_SYMBOL
        RegionalSetting_SHORT_DATE
        RegionalSetting_LONG_DATE
        RegionalSetting_CURRENCY_CODE
        RegionalSetting_COUNTRY
        RegionalSetting_THOUSAND_SEPARATOR
        RegionalSetting_TIME_SEPARATOR
        RegionalSetting_LIST_SEPARATOR
End Enum

' >< CONSTANTS ><
'
Private Const RDW_INVALIDATE                As Long = &H1

Private Const DEFAULT_CHARSET               As Long = 1
Private Const SYMBOL_CHARSET                As Long = 2
Private Const SHIFTJIS_CHARSET              As Long = 128
Private Const HANGEUL_CHARSET               As Long = 129
Private Const CHINESEBIG5_CHARSET           As Long = 136
Private Const CHINESESIMPLIFIED_CHARSET     As Long = 134
Private Const HEBREW_CHARSET                As Long = 177

Private Const CLR_INVALID                   As Long = &HFFFF

Public Const HH_DISPLAY_TOPIC               As Long = &H0
Public Const HH_HELP_CONTEXT                As Long = &HF   ' Display mapped numeric value in  dwData.

Private Const GUID_OK                      As Long = 0

' >< UDT ><
'
Private Type GUID
        Guid1                       As Long
        Guid2                       As Integer
        Guid3                       As Integer
        Guid4(0 To 7)               As Byte
End Type

' >< API ><
'
Private Declare PtrSafe Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal HWND As LongPtr) As Long
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal HWND As LongPtr, lprcUpdate As Any, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private Declare PtrSafe Function FindWindowStr Lib "user32" Alias "FindWindowA" (ByVal ClassName As String, ByVal WndName As String) As LongPtr
Private Declare PtrSafe Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare PtrSafe Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function OleTranslateColor Lib "OLEAUT32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As LongPtr, pccolorref As Long) As Long
Private Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "OLE32.DLL" (pGuid As GUID, ByVal PointerToString As LongPtr, ByVal MaxLength As Long) As Long
Public Declare PtrSafe Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As LongPtr, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As LongPtr) As LongPtr
'

Public Sub g_SetAccelerators(Frm As UserForm, Optional FontSize As Single = 8)
            
        Dim objCtl                  As Control
        Dim objPage                 As Page
        Dim strCaption              As String
    
On Error Resume Next
    
        For Each objCtl In Frm.Controls
                
                If TypeOf objCtl Is MultiPage Then
                        
                        For Each objPage In objCtl.Pages
                                                                        
                                strCaption = objPage.Caption
                
                                If (Err.Number = 0) Then
                                        Call g_SetCaption(objPage)
                                Else
                                        Err.Clear
                                End If
                        
                        Next objPage
                
                Else
                                        
                        strCaption = objCtl.Caption
                                        
                        If (Err.Number = 0) Then
                                Call g_SetCaption(objCtl)
                        Else
                                Err.Clear
                        End If
        
                        If (Err.Number <> 0) Then Err.Clear
                                            
                        objCtl.SelectionMargin = False
                
                End If

                If (Err.Number <> 0) Then Err.Clear
                
                Call g_SetProperFont(objCtl.Font, FontSize)

        Next objCtl
    
Controlled_Exit:

Exit Sub
    
End Sub

Public Sub g_SetCaption(Obj As Object)
        
        Dim intAcc                  As Integer
        Dim strAcc                  As String
        Dim strRet                  As String
                
        strRet = Obj.Caption
        
        intAcc = InStr(strRet, "&")

        If (intAcc > 0) Then

                strAcc = Mid$(strRet, (intAcc + 1), 1)
                strRet = Left$(strRet, (intAcc - 1)) & Right$(strRet, (Len(strRet) - intAcc))

                If (strAcc <> "&") Then Obj.Accelerator = strAcc

        End If
        
        Obj.Caption = strRet
                
        If (Err.Number <> 0) Then Err.Clear
        
End Sub

Public Sub g_SetProperFont(Obj As Object, ByVal FontSize As Single)

On Error GoTo ErrTrap
    
        Select Case GetUserDefaultLCID
                Case &H404 ' Traditional Chinese
                        Obj.Charset = CHINESEBIG5_CHARSET
                        Obj.name = ChrW$(&H65B0) + ChrW$(&H7D30) + ChrW$(&H660E) _
                         + ChrW$(&H9AD4)   'New Ming-Li
                        Obj.Size = 9
                Case &H411 ' Japan
                        Obj.Charset = SHIFTJIS_CHARSET
                        Obj.name = ChrW$(&HFF2D) + ChrW$(&HFF33) + ChrW$(&H20) + _
                                   ChrW$(&HFF30) + ChrW$(&H30B4) + ChrW$(&H30B7) + ChrW$(&H30C3) + _
                                   ChrW$(&H30AF)
                        Obj.Size = 9
                Case &H412 'Korea UserLCID
                        Obj.Charset = HANGEUL_CHARSET
                        Obj.name = ChrW$(&HAD74) + ChrW$(&HB9BC)
                        Obj.Size = 9
                Case &H804 ' Simplified Chinese
                        Obj.Charset = CHINESESIMPLIFIED_CHARSET
                        Obj.name = ChrW$(&H5B8B) + ChrW$(&H4F53)
                        Obj.Size = 9
                Case &H40D  ' Hebrew
                        Obj.Charset = HEBREW_CHARSET
                        Obj.Size = FontSize
                Case Else   ' The other countries
'                        'Do Nothing, makes Combo boxes crash
'                        obj.Charset = DEFAULT_CHARSET
'                        obj.Name = ""   ' Get the default UI font.
                        Obj.Size = FontSize
            End Select

Exit Sub
    
ErrTrap:
        
        Err.Number = Err
    
End Sub

Public Function gs_GetRegionalSetting(ByVal lSetting As SystemRegionalSetting) As String
    
        Dim strSecName              As String
        Dim strKey                  As String
        Dim strRetString            As String * 256
        Dim lngSuccess              As Long
        Dim strRet                  As String
    
On Error GoTo ErrTrap
        
        Select Case lSetting
                Case RegionalSetting_DATE_SEPARATOR: strKey = "sDate"
                Case RegionalSetting_DECIMAL_SYMBOL: strKey = "sDecimal"
                Case RegionalSetting_SHORT_DATE: strKey = "sShortDate"
                Case RegionalSetting_LONG_DATE: strKey = "sLongDate"
                Case RegionalSetting_CURRENCY_CODE: strKey = "sCurrency"
                Case RegionalSetting_COUNTRY: strKey = "sCountry"
                Case RegionalSetting_THOUSAND_SEPARATOR: strKey = "sThousand"
                Case RegionalSetting_TIME_SEPARATOR: strKey = "sTime"
                Case RegionalSetting_LIST_SEPARATOR: strKey = "sList"
        End Select
        
        strRet = vbNullString
        
        strSecName = "Intl"
        
        lngSuccess = GetProfileString(strSecName, strKey, vbNullString, strRetString, Len(strRetString))
        
        If (lngSuccess <> 0) Then
                strRet = Left$(strRetString, InStr(strRetString, Chr$(0)) - 1)
        End If

Controlled_Exit:
        
        gs_GetRegionalSetting = strRet

Exit Function

ErrTrap:
        
        MsgBox Err.Description, vbExclamation
        strRet = vbNullString
        Resume Controlled_Exit
        
End Function

Public Sub g_UnLoadAllForms(Optional ByVal bEnd As Boolean = False)

        Dim Frm                     As UserForm
    
On Error Resume Next
        
        ' loop thru all the forms in the project
        For Each Frm In VBA.UserForms
        
                Frm.Hide
                Call Unload(Frm)
                Set Frm = Nothing
               
                If (Err.Number <> 0) Then Err.Clear
                
        Next Frm
        
        App.ActiveDrawing.Redraw
        DoEvents

        ' hard end if needed
        If bEnd Then End

End Sub

Public Sub g_Eval(oTB As TextBox, bCancel As MSForms.ReturnBoolean, _
                  Optional ByVal bAllowNeg As Boolean = True, Optional ByVal bUseTag As Boolean = False)
                    
        Dim dblX                    As Double
        Dim strX                    As String

On Error GoTo ErrTrap

        ' start out cool
        bCancel = False
    
        With oTB
    
                If (Len(Trim$(.Text)) = 0) Then Exit Sub
                If (.SelLength > 0) Then Exit Sub
        
                dblX = App.Frame.Evaluate(gs_NoComma(.Text))
                
                If Not bAllowNeg Then dblX = Abs(dblX)
                
                strX = Format$(CStr(dblX), "#0.0000")
                
                ' 07/15/08 - rg
                '
                If bUseTag Then .Tag = dblX
                
                .Text = gs_NoZeros(gs_NoComma(strX))
                .SelStart = 0
    
        End With
    
Controlled_Exit:

Exit Sub

ErrTrap:

        MsgBox Err.Description, vbInformation
        bCancel = True
        
        With oTB
                .SetFocus
                .SelStart = 0
        End With
    
        Resume Controlled_Exit

End Sub

Public Sub g_EvalInt(oTB As TextBox, bCancel As MSForms.ReturnBoolean, ByVal bAllowNeg As Boolean)
        
        Call g_Eval(oTB, bCancel)
        
        ' if ok then convert to integer
        If Not bCancel Then
                With oTB
                        If bAllowNeg Then
                                .Text = CInt(.Text)
                        Else
                                .Text = Abs(CInt(.Text))
                        End If
                End With
        End If

Controlled_Exit:

Exit Sub

End Sub

Public Sub g_EvalS(sVal As String, sReturn As String, Cancel As Boolean)

        Dim X                       As Double
    
        If (Len(Trim$(sVal)) = 0) Then sVal = "0"

On Error GoTo ErrTrap

        X = App.Frame.Evaluate(gs_NoComma(sVal))
        sReturn = Format$(X, "#0.0###")
        sReturn = gs_NoZeros(sReturn)

Exit Sub

ErrTrap:

        Cancel = True
        MsgBox Err.Description, vbExclamation
        
End Sub

Public Sub g_EvalSLng(sVal As String, sReturn As String, Cancel As Boolean)

        Dim X                       As Double
    
        If (Len(Trim$(sVal)) = 0) Then sVal = "0"

On Error GoTo ErrTrap

        X = App.Frame.Evaluate(gs_NoComma(sVal))
        sReturn = CStr(CLng(X))
        
Exit Sub

ErrTrap:

        Cancel = True
        MsgBox Err.Description, vbExclamation
        
End Sub

Public Function gb_IsValOK(ByVal sVal As String, Optional ByVal bDecimal As Boolean = False, Optional ByVal bNegative As Boolean = False) As Boolean

        Dim strValidChars           As String
        Dim strChar                 As String
        Dim I                       As Integer
        Dim intMax                  As Integer
        Dim intDec                  As Integer
        Dim intNeg                  As Integer

        ' assign chars
        strValidChars = "0123456789"
        intNeg = 0
        intDec = 0

        ' assume success
        gb_IsValOK = True
        
        ' allow decimal?
        If bDecimal Then strValidChars = strValidChars & "."
           
        ' allow negative
        If bNegative Then strValidChars = strValidChars & "-"
           
        intMax = Len(sVal)
        
        ' loop string
        For I = 1 To intMax
            
                ' get a char
                strChar = Mid$(sVal, I, 1)
               
                If strChar = "." Then intDec = (intDec + 1)
                If strChar = "-" Then intNeg = (intNeg + 1)
                
                Select Case True
                
                        Case (InStr(1, strValidChars, strChar) = 0), (intDec > 1), (intNeg > 1)
                                gb_IsValOK = False
                                Exit Function
                End Select
            
        Next I

End Function

Public Function gb_CheckAllText(Container As Object) As Boolean
        
        Dim Ctl                     As Control
        
        'start out cool
        gb_CheckAllText = True
        
        'check for empty text boxes and warn user if any
        For Each Ctl In Container.Controls
        
                If TypeOf Ctl Is TextBox Then
                        
                        If (Ctl.Tag <> "999") Then
                        
                                ' check for empty if enabled
                                If Ctl.Enabled Then
                                        If (Len(Trim$(Ctl.Text)) = 0) Then
                                                MsgBox "Please complete all information." & Space$(3), vbInformation, App.name
                                                Ctl.SetFocus
                                                gb_CheckAllText = False
                                                Exit For
                                        End If
                                End If
                        
                        End If
                        
                End If
                
        Next Ctl
        
        Set Ctl = Nothing

Exit Function

End Function

Public Function gcol_DelimitedStringToCollection(ByVal sSubItems As String, Optional ByVal sDelimitChar As String = ",", Optional ByVal bIncludeNulls As Boolean = False) As Collection

        Dim strLen                  As Long
        Dim I                       As Long
        Dim J                       As Long
        Dim lngDelLen               As Long
        Dim strTemp                 As String
        Dim strRet                  As String
        Dim colRet                  As Collection
        
        ' Delimits a string, using the specified character(s) or the default
        ' comma, and returns a collection of strings
        
        Set colRet = New Collection
  
        If (sSubItems = vbNullString) Then
                If bIncludeNulls Then
                        Call colRet.add(vbNullString)
                        GoTo Controlled_Exit
                End If
        End If
        
        lngDelLen = Len(sDelimitChar)
        
        If (Right$(sSubItems, 1) <> sDelimitChar) Then
                strTemp = sSubItems & sDelimitChar
        Else
                strTemp = sSubItems
        End If
  
        strLen = Len(strTemp)
  
        J = 1
        I = 1
  
        While (J < strLen) And (J > 0)
                
                J = InStr(I, strTemp, sDelimitChar)
                
                If ((I * J) > 0) Then
      
                        strRet = Trim(Mid$(strTemp, I, J - I))
                        
                        If (strRet <> vbNullString) Then
                                Call colRet.add(strRet)
                        Else
                                If (strRet = vbNullString) Then
                                        If bIncludeNulls Then Call colRet.add(strRet)
                                End If
                        End If
                        
                End If
                
                I = (J + lngDelLen)
                
        Wend
        
Controlled_Exit:
        
        Set gcol_DelimitedStringToCollection = colRet
        
        Set colRet = Nothing
        
Exit Function
  
End Function

Public Function gv_Split(ByVal sString As String, Optional ByVal sDelimeter As String = ",", Optional ByVal bBase1 As Boolean = False) As Variant

        Dim strSDelim               As String
        Dim strString               As String
        Dim intIStringLength        As Integer
        Dim intIDelimPosition       As Integer
        Dim strSDoubleQuoteMark     As String
        Dim intIIndex               As Integer
        Dim arystrAData1()          As String
        Dim strSDatafield           As String
        
        strString = sString
        strSDelim = sDelimeter
        strSDoubleQuoteMark = Chr$(34)
        intIStringLength = Len(strString)
        
        intIIndex = IIf(bBase1, 1, 0)
        
        ' if the length of the data string is greater than zero
        If (intIStringLength > 0) Then
        
                'Debug.Print strString
        
                ' search for a sDelimiter in the datastring
                intIDelimPosition = InStr(strString, strSDelim)
                
                Do While (intIDelimPosition <> 0)
                
                        ' snag the datafield
                        strSDatafield = Trim$(Left$(strString, (intIDelimPosition - 1)))
                        
                        ' look for and remove leading/trailing quotes,
                        ' leave if only one or the other or within the string
                        If (Left$(strSDatafield, 1) = strSDoubleQuoteMark) Then
                                
                                If (Right$(strSDatafield, 1) = strSDoubleQuoteMark) Then
                                        strSDatafield = Left$(strSDatafield, (Len(strSDatafield) - 1))
                                        strSDatafield = Right$(strSDatafield, (Len(strSDatafield) - 1))
                                End If
                        
                        End If

                        ' sort out the rest of the string
                        strString = Right$(strString, (Len(strString) - intIDelimPosition))
                                                                
                        ReDim Preserve arystrAData1(intIIndex)
                        arystrAData1(intIIndex) = strSDatafield
                        intIDelimPosition = InStr(strString, strSDelim)
                        
                        intIIndex = (intIIndex + 1)
                        
                Loop
                
                'iIndex = iIndex + 1
                ReDim Preserve arystrAData1(intIIndex)
                arystrAData1(intIIndex) = strString
                
        End If
        
        gv_Split = arystrAData1
        
End Function

Public Function gs_NoComma(ByVal sVal As String) As String

        ' this function is designed to replace the comma in the German
        ' numbering system with a decimal point as needed by Alphacam

        Dim strRet                  As String
        
On Error GoTo ErrTrap
        
        strRet = Replace$(sVal, ",", ".")
        
Controlled_Exit:
        
        gs_NoComma = strRet

Exit Function

ErrTrap:
        
        strRet = sVal
        Resume Controlled_Exit

End Function

Public Function gs_ReplaceSpaces(ByVal sVal As String, Optional ByVal sChr As String = "_")
        
        Dim strRet                  As String
        
        strRet = Replace$(sVal, Space$(1), sChr)
        
        gs_ReplaceSpaces = strRet

End Function

Public Function gs_RemoveIllegalChars(ByVal sVal As String, ByVal bReplaceDecWithP As Boolean, Optional ByVal sChr As String = "_") As String

        Dim I                       As Integer
        Dim intMax                  As Integer
        Dim strChar                 As String
        Dim strRet                  As String

On Error GoTo ErrTrap
        
        Const DEF_ILLEGAL           As String = "\/:*?<>|"""
        
        ' set default return val
        strRet = vbNullString
        
        sVal = Trim$(sVal)
        intMax = Len(sVal)
                
        If (intMax > 0) Then
        
                For I = 1 To intMax
                
                        strChar = Mid$(sVal, I, 1)
                        
                        If InStr(DEF_ILLEGAL, strChar) > 0 Then
                                strChar = sChr
                                strRet = strRet & strChar
                        Else
                                If bReplaceDecWithP Then
                                        If (strChar = ".") Then
                                                strRet = strRet & "P"
                                        Else
                                                strRet = strRet & strChar
                                        End If
                                Else
                                        strRet = strRet & strChar
                                End If
                        End If
                    
                Next I
        
        End If
                                    
Controlled_Exit:

        gs_RemoveIllegalChars = strRet

Exit Function

ErrTrap:
    
        MsgBox Err.Description, vbExclamation
        strRet = sVal
        Resume Controlled_Exit
    
End Function

Public Function gs_RemoveNullChars(ByVal sVal As String) As String
        gs_RemoveNullChars = Replace$(sVal, Chr$(0), vbNullString)
End Function

Public Function gs_DateToString(ByVal dtNow As Date) As String
        gs_DateToString = Trim$(Str$(Year(dtNow))) + Trim$(Str$(Month(dtNow))) + Trim$(Str$(Day(dtNow)))
End Function

Public Function gs_TruncateText(ByVal sText As String, ByVal lControlWidth As Long) As String
        
        Dim strText                 As String
        Dim strLeft                 As String
        Dim strRight                As String
        Dim lngMid                  As Long
        Dim lngLen                  As Long
        Dim lngTrimR                As Long
        Dim lngTrimL                As Long
        Dim lngBackSlash            As Long
        Dim lngMax                  As Long
        
        Const DEF_ELLIPSE           As String = "..."
        
On Error Resume Next
        
        ' init
        gs_TruncateText = sText
                        
        ' get the overall length and the middle char
        strText = sText
        lngLen = Len(strText)
        lngMax = ((lControlWidth * 0.25) - 7)
        
        If (lngLen >= lngMax) Then
                                
                ' we'll look for a backslash in hopes to leave the entire
                ' file/folder name visible if tuncating a file/folder path
                lngBackSlash = InStrRev(strText, "\")

                If (lngBackSlash > 0) Then
                        lngMid = (lngBackSlash - 1)
                        lngTrimL = (lngLen - lngMax)
                        lngTrimR = 0
                Else
                        lngMid = (lngLen / 2)
                        lngTrimL = ((lngLen - lngMax) / 2)
                        lngTrimR = lngTrimL
                End If
                
                'Debug.Print "MID: " & lngMid & " ~ TRIM_L: " & lngTrimL & " ~ TRIM_R: " & lngTrimR
                
                strLeft = Left$(strText, lngMid)
                strRight = Right$(strText, (lngLen - lngMid))
                
                strLeft = Left$(strLeft, (Len(strLeft) - lngTrimL))
                strLeft = Left$(strLeft, (Len(strLeft) - 3))
                strRight = Right$(strRight, (Len(strRight) - lngTrimR))
                
                'Debug.Print "LEFT: " & strLeft & " ~ RIGHT: " & strRight
                
                strText = strLeft & DEF_ELLIPSE & strRight

        End If

        'Debug.Print "MAX: " & lMaxWidth & " ~ ACTUAL: " & Len(strText)
                
        gs_TruncateText = strText

Exit Function

End Function

Public Sub g_EnableDisableControls(oContainer As Object, ByVal bEnable As Boolean, ByVal bIncludeContainer As Boolean)
        
        Dim Ctl                     As Control
        
On Error Resume Next
        
        For Each Ctl In oContainer.Controls
                
                Ctl.Enabled = bEnable
                
                If (Err.Number <> 0) Then Err.Clear
        
        Next Ctl
        
        If bIncludeContainer Then oContainer.Enabled = bEnable
        If (Err.Number <> 0) Then Err.Clear
                        
        Set Ctl = Nothing

End Sub

Public Sub g_Repaint(Frm As UserForm)
        
        Dim lngRet                  As Long
        Dim lngHwnd                 As LongPtr
        
        lngHwnd = gl_FrmHwnd(Frm)
        
        If (IsWindow(lngHwnd) = 0) Then Exit Sub
        
        ' redraw the userform
        lngRet = RedrawWindow(lngHwnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE)
    
End Sub

Public Function gb_IsInCollection(C As Collection, ByVal V As Variant) As Boolean

On Error GoTo ErrTrap:
        
        ' simple check - will raise error if item does not exist in collection
        With C(V)
        End With
  
        gb_IsInCollection = True
  
Exit Function
  
ErrTrap:

        gb_IsInCollection = False
        
End Function

Public Sub g_GetArrayBounds(vArray As Variant, lL As Long, lU As Long)
                
        Dim lngL                    As Long
        Dim lngU                    As Long
        
On Error GoTo ErrTrap
        
        lngL = LBound(vArray)
        lngU = UBound(vArray)
        
Controlled_Exit:

        lL = lngL
        lU = lngU
        
Exit Sub

ErrTrap:
        
        Err.Clear
        lngL = 0
        lngU = -1                   ' set to -1 to avoid loop in for-next
        Resume Controlled_Exit

End Sub

Public Function gl_FrmHwnd(ByRef Frm As Object) As Variant
    
        Dim vRet                    As Variant

On Error GoTo ErrTrap
        
        ' Assume handle will not be found.
        vRet = 0
        
        ' First check for form under Visual Basic for
        ' Applications 6.0 or Visual Basic 5.0/6.0 IDEs.
        vRet = FindWindowStr("ThunderDFrame", Frm.Caption)
        
        ' If handle is not found then keep looking
        If (vRet = 0) Then
        
                ' Check for form under Visual Basic for Applications 5.0 IDE.
                vRet = FindWindowStr("ThunderXFrame", Frm.Caption)
                
                ' If handle is not found--
                If (vRet = 0) Then
                
                        ' Check for form compiled from MSForms
                        ' object library dated 3/22/99 or later.
                        vRet = FindWindowStr("ThunderRT6DFrame", Frm.Caption)
                        
                        ' If handle is not found--
                        If (vRet = 0) Then
                                ' Check for form compiled from initial
                                ' version of MSForms object library.
                                vRet = FindWindowStr("ThunderRT5DFrame", Frm.Caption)
                        End If
                        
                End If
                
        End If
   
ErrTrap:
        
        If Err Then vRet = CVErr(Err)
        gl_FrmHwnd = vRet
   
End Function

Public Function gs_NoZeros(ByVal sVal As String) As String
        
        Dim strVal                  As String
        
On Error Resume Next
        
        strVal = gs_NoComma(sVal)
        
        If (InStr(strVal, ".") > 0) Then
        
                Do While Right$(strVal, 1) = "0"
                        strVal = Left$(strVal, Len(strVal) - 1)
                Loop
                
                If Right$(strVal, 1) = "." Then
                        strVal = Left$(strVal, Len(strVal) - 1)
                End If
                
                If (Len(Trim$(strVal)) = 0) Then strVal = "0"
                
        End If

        gs_NoZeros = strVal
    
Controlled_Exit:

Exit Function
        
End Function

Public Function gs_RemovePointFromZero(ByVal sVal As String) As String
        
        Dim strRet                  As String
        
        strRet = sVal
        
        If (PSDbl(sVal) = 0) Then strRet = "0"
        
        gs_RemovePointFromZero = strRet
        
End Function

Public Function gs_StripCR(ByVal S As String) As String
        
        Dim dblLen                  As Double
        Dim strRet                  As String
        
        dblLen = Len(Trim$(S))
        strRet = S

        If CBool(dblLen) Then
                If (Right$(strRet, 1) = Chr$(10)) Then
                        strRet = Left$(strRet, (dblLen - 2))
                End If
        End If

        gs_StripCR = strRet

End Function

Public Function gs_StripLF(ByVal sVal As String) As String

        Dim sRet                    As String
        
        sRet = sVal
        
        Do While (Right$(sRet, 1) = Chr$(10))
                sRet = Left$(sRet, (Len(Trim$(sRet)) - 1))
        Loop

        gs_StripLF = sRet

End Function

Public Function gs_StripCRLF(ByVal S As String) As String

        Dim strRet                  As String
        
        strRet = S
        strRet = Trim$(Replace$(strRet, vbCr, vbNullString))
        strRet = Trim$(Replace$(strRet, vbLf, vbNullString))
        
        gs_StripCRLF = strRet

End Function

Public Function gs_GUID() As String
 
        Dim udtGUID                 As GUID
        Dim lngResult               As Long
        Dim strRet                  As String
  
        Const GUID_LENGTH           As Long = 38

        lngResult = CoCreateGuid(udtGUID)

        If (lngResult = GUID_OK) Then
                strRet = String$(GUID_LENGTH, 0)
                lngResult = StringFromGUID2(udtGUID, StrPtr(strRet), (GUID_LENGTH + 1))
        Else
                strRet = vbNullString
        End If
        
        gs_GUID = strRet
        
End Function

Public Function gl_BinaryValue(ByVal lNumber As Long, Optional ByVal lBinary As Long = 0) As Long

        Dim lngRet                  As Long
        Dim dblCheck                As Double

        Const DEF_MAX_VAL           As Long = 2147483647

On Error GoTo ErrTrap

        dblCheck = (2 ^ (lNumber - 1))

        ' prevent potential overflow
        Select Case True

                Case (dblCheck > DEF_MAX_VAL), _
                     ((dblCheck + lBinary) > DEF_MAX_VAL)

                        lngRet = 0

                Case Else

                        lngRet = CLng(dblCheck)

        End Select

Controlled_Exit:

        gl_BinaryValue = lngRet

Exit Function

ErrTrap:

        lngRet = 0
        Resume Controlled_Exit

End Function

Public Function gl_TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
        If OleTranslateColor(clr, hPal, gl_TranslateColor) Then gl_TranslateColor = CLR_INVALID
End Function

Public Sub g_Help(ByVal sCHM As String, Optional ByVal lIndex As Long = 0)
                
        Dim FSO                     As New Scripting.FileSystemObject
        Dim lngPtrRet               As LongPtr
        
        ' if invalid file, then bail
        If Not FSO.FileExists(sCHM) Then Exit Sub
        
        If (lIndex <> 0) Then
                
                ' try to launch context
                lngPtrRet = HtmlHelp(0, sCHM, HH_HELP_CONTEXT, lIndex)

                ' if context failed, then simply launch the main topic
                If (lngPtrRet = 0) Then
                        lngPtrRet = HtmlHelp(0, sCHM, HH_DISPLAY_TOPIC, 0)
                End If
                
        Else
                lngPtrRet = HtmlHelp(0, sCHM, HH_DISPLAY_TOPIC, 0)
        End If

End Sub

Public Sub g_DebugNote(ByVal sDebugString As String)
        ' outputs string to external debug viewer (e.g. DEBUGMON.exe)
        Call OutputDebugString(DEF_MACRO_NAME & ": " & sDebugString & vbCrLf)
End Sub
-------------------------------------------------------------------------------
