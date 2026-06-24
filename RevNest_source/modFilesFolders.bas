in file: C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\000\StartUp\Utils\ReverseNest\ReverseNest.amb - OLE stream: 'vao/The VBA Project/_VBA_Project/VBA/modFilesFolders'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module

' >< API ><
'
Private Declare PtrSafe Sub OleInitialize Lib "OLE32.DLL" (pvReserved As Any)
Private Declare PtrSafe Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal hMem As LongPtr)
Private Declare PtrSafe Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As LongPtr, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BROWSEINFO) As Long
Private Declare PtrSafe Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare PtrSafe Function SHParseDisplayName Lib "shell32.dll" (ByVal pszName As LongPtr, ByVal pbc As LongPtr, ByRef ppidl As Long, ByVal sfgaoIn As Long, ByRef psfgaoOut As LongPtr) As Long
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As LongPtr, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal HWND As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal HWND As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal HWND As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As LongPtr, ByVal nFolder As Long, ByVal hToken As LongPtr, ByVal dwFlags As Long, ByVal pszPath As String) As LongPtr

' >< UDT ><
'
Private Type RECT
        Left                        As Long
        Top                         As Long
        Right                       As Long
        Bottom                      As Long
End Type

Private Type SH_ITEM_ID
        cb                              As Long
        abID                            As Byte
End Type

Private Type ITEMIDLIST
        mkid                            As SH_ITEM_ID
End Type

Private Type BROWSEINFO
        hwndOwner                       As LongPtr
        pIDLRoot                        As Long
        pszDisplayName                  As String
        lpszTitle                       As String
        ulFlags                         As Long
        lpfnCallback                    As LongPtr
        lParam                          As LongPtr
        iImage                          As Long
End Type

' >< ENUMS ><
'
Public Enum SystemSpecialFolder
        sysSpecialFolder_AdministrativeTools = &H30&
        sysSpecialFolder_CommonAdministrativeTools = &H2F&
        sysSpecialFolder_ApplicationData = &H1A&
        sysSpecialFolder_CommonAppData = &H23&
        sysSpecialFolder_CommonDocuments = &H2E&
        sysSpecialFolder_Cookies = &H21&
        sysSpecialFolder_History = &H22&
        sysSpecialFolder_InternetCache = &H20&
        sysSpecialFolder_LocalApplicationData = &H1C&
        sysSpecialFolder_MyPictures = &H27&
        sysSpecialFolder_Personal = &H5&
        sysSpecialFolder_ProgramFiles = &H26&
        sysSpecialFolder_CommonProgramFiles = &H2B&
        sysSpecialFolder_System = &H25&
        sysSpecialFolder_Windows = &H24&
        sysSpecialFolder_Fonts = &H14&
End Enum

' >< CONSTANTS ><
'
Private Const WM_INITDIALOG                 As Long = &H110

Private Const WM_USER                       As Long = &H400
Private Const BFFM_INITIALIZED              As Long = 1
Private Const BFFM_SELCHANGED               As Long = 2
Private Const BFFM_SETSTATUSTEXT            As Long = (WM_USER + 100)
Private Const BFFM_SETSELECTION             As Long = (WM_USER + 102)
Private Const BIF_DEFAULT                   As Long = &H0
Private Const BIF_RETURNONLYFSDIRS          As Long = &H1       ' only local Directory
Private Const BIF_DONTGOBELOWDOMAIN         As Long = &H2
Private Const BIF_STATUSTEXT                As Long = &H4       ' Not With BIF_NEWDIALOGSTYLE
Private Const BIF_RETURNFSANCESTORS         As Long = &H8
Private Const BIF_EDITBOX                   As Long = &H10
Private Const BIF_VALIDATE                  As Long = &H20      ' use With BIF_EDITBOX or BIF_USENEWUI
Private Const BIF_NEWDIALOGSTYLE            As Long = &H40      ' Use OleInitialize before
Private Const BIF_USENEWUI                  As Long = &H50      ' = (BIF_NEWDIALOGSTYLE + BIF_EDITBOX)
Private Const BIF_BROWSEINCLUDEURLS         As Long = &H80
Private Const BIF_UAHINT                    As Long = &H100     ' use With BIF_NEWDIALOGSTYLE, add Usage Hint if no EditBox
Private Const BIF_NONEWFOLDERBUTTON         As Long = &H200
Private Const BIF_NOTRANSLATETARGETS        As Long = &H400
Private Const BIF_BROWSEFORCOMPUTER         As Long = &H1000
Private Const BIF_BROWSEFORPRINTER          As Long = &H2000
Private Const BIF_BROWSEINCLUDEFILES        As Long = &H4000
Private Const BIF_SHAREABLE                 As Long = &H8000    ' use With BIF_NEWDIALOGSTYLE

Private Const MAX_PATH                      As Long = 260

Private m_sCurrentDir                       As String
'

Public Function gs_ReadFileContents(ByVal sFile As String) As String
        
        Dim FSO                     As Scripting.FileSystemObject
        Dim strRet                  As String
        
On Error GoTo ErrTrap
        
        Set FSO = New Scripting.FileSystemObject
        
        strRet = vbNullString
        
        If FSO.FileExists(sFile) Then
                strRet = FSO.GetFile(sFile).OpenAsTextStream.ReadAll
        End If
        
Controlled_Exit:

        gs_ReadFileContents = strRet
        
        Set FSO = Nothing
        
Exit Function

ErrTrap:
        
        strRet = vbNullString
        MsgBox Err.Description, vbExclamation, App.name
        Resume Controlled_Exit
        
End Function

Public Function gs_EnsureBackslash(ByVal sPath As String) As String
        
        Dim strRet                  As String
        
        strRet = sPath
        
        If (Right$(sPath, 1) <> "\") Then strRet = sPath & "\"
        
        gs_EnsureBackslash = strRet

End Function

Public Function gs_StripLeadingBackslash(ByVal sPath As String) As String
        
        Dim strRet                  As String
        
        strRet = sPath
        
        Do While (Left$(strRet, 1) = "\")
                strRet = Right$(strRet, (Len(strRet) - 1))
        Loop
        
        gs_StripLeadingBackslash = strRet

End Function

Public Function gs_ParseFileName(ByVal sPath As String, ByVal bStripExtension As Boolean) As String
  
        Dim strRet                  As String
        Dim intX                    As Integer
        
        intX = InStrRev(sPath, "\")
        
        strRet = Trim$(Right$(sPath, Len(sPath) - intX))
        
        If bStripExtension Then strRet = gs_StripFileExtension(strRet)

        If (Right$(strRet, 1) = Chr$(0)) Then
                strRet = Left$(strRet, (Len(strRet) - 1))
        End If

Controlled_Exit:
        
        gs_ParseFileName = strRet
        
Exit Function

End Function

Public Function gs_ParseDirName(ByVal sPath As String, ByVal bIncludeBackslash As Boolean) As String
  
        Dim strRet                  As String
        Dim intX                    As Integer
        
        strRet = Trim$(sPath)
        
        If (Len(Trim$(strRet)) > 0) Then
        
              intX = InStrRev(sPath, "\")
              
              strRet = Trim$(Left$(sPath, (intX - 1)))
              
              If Left$(strRet, 1) = Chr$(0) Then
                      strRet = Left$(strRet, (Len(strRet) - 1))
              End If
        
              If bIncludeBackslash Then strRet = gs_EnsureBackslash(strRet)
        
        End If
        
Controlled_Exit:

        gs_ParseDirName = strRet

Exit Function

End Function

Public Function gs_ParseFileExtension(ByVal sFullPath As String, ByVal bIncludePoint As Boolean) As String
    
        Dim intPoint                As Integer
        Dim strRet                  As String
        
        strRet = vbNullString
        
        intPoint = InStrRev(sFullPath, ".")
        
        If (intPoint > 0) Then
        
                If bIncludePoint Then
                        strRet = UCase$(Right$(sFullPath, ((Len(sFullPath) - intPoint) + 1)))
                Else
                        strRet = UCase$(Right$(sFullPath, (Len(sFullPath) - intPoint)))
                End If
                
        End If
        
Controlled_Exit:

        gs_ParseFileExtension = strRet

Exit Function

End Function

Public Function gs_StripFileExtension(ByVal sFile As String) As String
        
        Dim strRet                  As String
        Dim intPoint                As Integer
        
        ' set default return val
        strRet = sFile
        
        intPoint = InStrRev(sFile, ".")
        
        If (intPoint > 0) Then
                strRet = Left$(sFile, (Len(sFile) - ((Len(sFile) - intPoint) + 1)))
        End If
        
Controlled_Exit:

        gs_StripFileExtension = strRet

Exit Function

End Function

Public Function gs_ReplaceFileExtension(ByVal sFile As String, ByVal sNewExt As String) As String
    
        Dim strRet                  As String
        
        strRet = gs_StripFileExtension(sFile)
        
        ' only add the new file extension if there is one
        If (Len(Trim$(sNewExt)) > 0) Then
                If (Left$(sNewExt, 1) = ".") Then
                        strRet = strRet & sNewExt
                Else
                        strRet = strRet & "." & sNewExt
                End If
        End If
               
Controlled_Exit:

        gs_ReplaceFileExtension = strRet

Exit Function

End Function

Public Function gs_GetLocalAppDataDir(Optional ByVal bIncludeBackslash As Boolean = True) As String
        
        Dim FSO                     As Scripting.FileSystemObject
        Dim strPath                 As String
        Dim strRet                  As String
        Dim blnRet                  As Boolean

        ' this function will return the path to the users application data dir specific to this addin
        '
        ' e.g...
        '
        ' C:\Users\<USER_NAME>\AppData\Local\Planit\Alphacam\<NAME_OF_ADDIN>\
        
        blnRet = False
        
        Set FSO = New Scripting.FileSystemObject
        
        strRet = gs_GetSpecialFolder(sysSpecialFolder_LocalApplicationData, True)
        
        ' first, lets see if the dir is already there
        strPath = strRet & "Planit\Alphacam\" & DEF_MACRO_NAME
        '
        If FSO.FolderExists(strPath) Then
                strRet = strPath
                blnRet = True
                GoTo Controlled_Exit
        End If
        
        If FSO.FolderExists(strRet) Then
                
                strRet = strRet & "Planit"
                
                If gb_EnsureDirExistance(strRet) Then
                        
                        strRet = strRet & "Alphacam"
                        
                        If gb_EnsureDirExistance(strRet) Then
                                
                                strRet = strRet & DEF_MACRO_NAME
                                
                                If gb_EnsureDirExistance(strRet) Then
                                        blnRet = True
                                End If
                        
                        End If
                        
                End If
                
        End If
        
Controlled_Exit:

        If blnRet Then
                If bIncludeBackslash Then strRet = gs_EnsureBackslash(strRet)
        End If
        
        gs_GetLocalAppDataDir = strRet
        
        Set FSO = Nothing

Exit Function
        
End Function

Public Function gs_GetCommonAppDataDir(Optional ByVal bIncludeBackslash As Boolean = True) As String
        
        Dim FSO                     As Scripting.FileSystemObject
        Dim strPath                 As String
        Dim strRet                  As String
        Dim blnRet                  As Boolean

        ' this function will return the path to the common application data dir specific to this addin
        '
        ' e.g...
        '
        ' C:\ProgramData\Planit\Alphacam\<NAME_OF_ADDIN>\
        
        blnRet = False
        
        Set FSO = New Scripting.FileSystemObject
        
        strRet = gs_GetSpecialFolder(sysSpecialFolder_CommonAppData, True)
        
        ' first, lets see if the dir is already there
        strPath = strRet & "Planit\Alphacam\" & DEF_MACRO_NAME
        '
        If FSO.FolderExists(strPath) Then
                strRet = strPath
                blnRet = True
                GoTo Controlled_Exit
        End If
        
        If FSO.FolderExists(strRet) Then
                
                strRet = strRet & "Planit"
                
                If gb_EnsureDirExistance(strRet) Then
                        
                        strRet = strRet & "Alphacam"
                        
                        If gb_EnsureDirExistance(strRet) Then
                                
                                strRet = strRet & DEF_MACRO_NAME
                                
                                If gb_EnsureDirExistance(strRet) Then
                                        blnRet = True
                                End If
                        
                        End If
                        
                End If
                
        End If
        
Controlled_Exit:

        If blnRet Then
                If bIncludeBackslash Then strRet = gs_EnsureBackslash(strRet)
        End If
        
        gs_GetCommonAppDataDir = strRet
        
        Set FSO = Nothing

Exit Function
        
End Function

Public Function gs_GetSpecialFolder(ByVal iSpecialFolder As SystemSpecialFolder, _
                                    Optional ByVal bCreate As Boolean = False, _
                                    Optional ByVal bIncludeBackslash As Boolean = True) As String
        
        Dim lngPtrRet               As LongPtr
        Dim lngFlag                 As Long
        Dim strRet                  As String
        
        Const SHGFP_TYPE_CURRENT    As Long = 0
        Const SHGFP_TYPE_DEFAULT    As Long = 1
        Const MAX_PATH              As Long = 260
        Const CSIDL_FLAG_CREATE     As Long = &H8000&
        Const S_OK                  As Long = &H0           ' Success
        'Const S_FALSE               As Long = &H1           ' The Folder is valid, but does not exist
        'Const E_INVALIDARG          As Long = &H80070057    ' Invalid CSIDL Value
                
        strRet = String$(MAX_PATH, 0)
        
        If bCreate Then
                lngPtrRet = SHGetFolderPath(0, iSpecialFolder Or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, strRet)
        Else
                lngPtrRet = SHGetFolderPath(0, iSpecialFolder, 0, SHGFP_TYPE_CURRENT, strRet)
        End If
        
        If (lngPtrRet = S_OK) Then
                
                ' return the string upto the first null character
                strRet = Left$(strRet, InStr(1, strRet, Chr(0)) - 1)
        
                If bIncludeBackslash Then strRet = gs_EnsureBackslash(strRet)
                            
        End If
                
        gs_GetSpecialFolder = strRet
        
End Function

Public Function gs_FileSize(ByVal sFile As String) As String
    
        Dim FSO                     As New Scripting.FileSystemObject
        Dim fsoFile                 As Scripting.File
        Dim lngBytes                As Long
        
        Const KB                    As Long = 1024
        Const MB                    As Long = 1024 * KB
        Const GB                    As Long = 1024 * MB
        
On Error Resume Next
        
        gs_FileSize = "0 bytes"
        
        ' if no file then bail
        If Not FSO.FileExists(sFile) Then Exit Function
        
        Set fsoFile = FSO.GetFile(FSO.GetFile(sFile))
        
        lngBytes = fsoFile.Size
        
        ' format the number
        Select Case True
        
                Case (lngBytes < KB): gs_FileSize = Format$(lngBytes) & " bytes"
                Case (lngBytes < MB): gs_FileSize = Format$(lngBytes / KB, "0.00") & " KB"
                Case (lngBytes < GB): gs_FileSize = Format$(lngBytes / MB, "0.00") & " MB"
                Case Else: gs_FileSize = Format$(lngBytes / GB, "0.00") & " GB"
            
        End Select
        
Controlled_Exit:
        
        Set fsoFile = Nothing
        Set FSO = Nothing

Exit Function

End Function

Public Function gs_GetDir(Optional ByVal sTitle As String = "", Optional ByVal sRootDir As String = "", _
                          Optional ByVal sStartDir As String = "", Optional lOwnerHwnd As Long = 0, _
                          Optional ByVal bAllowNew As Boolean = True, _
                          Optional ByVal bIncludeFiles As Boolean = False, _
                          Optional ByVal bOnlyMyComputer As Boolean = False) As String

        Dim lngIDList               As Long
        Dim lngIDList2              As Long
        Dim IDL                     As ITEMIDLIST
        Dim strBuffer               As String
        Dim BI                      As BROWSEINFO
        Dim lngRet                  As Long
        Dim strRet                  As String
                
        strRet = vbNullString
        
        If (Len(sRootDir) > 0) Then
        
                If PathIsDirectory(sRootDir) Then
                        lngRet = SHParseDisplayName(StrPtr(sRootDir), ByVal 0&, lngIDList2, ByVal 0&, ByVal 0&)
                        BI.pIDLRoot = lngIDList2
                Else
                        If bOnlyMyComputer Then
                                lngRet = SHGetSpecialFolderLocation(ByVal 0&, &H11, IDL)  '= Start @ "My Computer" Folder
                        Else
                                lngRet = SHGetSpecialFolderLocation(ByVal 0&, 0&, IDL)  ' = Start @ "Desktop" Folder
                        End If
                        If (lngRet = 0) Then BI.pIDLRoot = IDL.mkid.cb
                End If

        Else
                If bOnlyMyComputer Then
                        lngRet = SHGetSpecialFolderLocation(ByVal 0&, &H11, IDL)  '= Start @ "My Computer" Folder
                Else
                        lngRet = SHGetSpecialFolderLocation(ByVal 0&, 0&, IDL)  ' = Start @ "Desktop" Folder
                End If
                If (lngRet = 0) Then BI.pIDLRoot = IDL.mkid.cb
        End If

        If (Len(sStartDir) > 0) Then
                m_sCurrentDir = sStartDir & vbNullChar
        Else
                m_sCurrentDir = vbNullChar
        End If
        
        With BI
        
                If (Len(sTitle) > 0) Then
                        .lpszTitle = sTitle
                Else
                        .lpszTitle = "Select A Directory"
                End If
        
                .lpfnCallback = ml_GetAddressofFunction(AddressOf ml_BrowseCallbackProc)
                .ulFlags = BIF_RETURNONLYFSDIRS
                If bIncludeFiles Then .ulFlags = .ulFlags + BIF_BROWSEINCLUDEFILES
        
                If bAllowNew Then
                        .ulFlags = .ulFlags + BIF_NEWDIALOGSTYLE + BIF_UAHINT
                        Call OleInitialize(Null) ' Initialize OLE and COM
                Else
                        .ulFlags = .ulFlags + BIF_STATUSTEXT
                End If
        
                If (lOwnerHwnd <> 0) Then .hwndOwner = lOwnerHwnd
        
        End With
        
        lngIDList = SHBrowseForFolder(BI)

        If (Len(sRootDir) > 0) Then
                If PathIsDirectory(sRootDir) Then Call CoTaskMemFree(lngIDList2)
        End If

        If (lngIDList) Then
                strBuffer = Space$(MAX_PATH)
                lngRet = SHGetPathFromIDList(lngIDList, strBuffer)
                Call CoTaskMemFree(lngIDList)
                strBuffer = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
                strRet = strBuffer
        Else
                strRet = ""
        End If
        
        gs_GetDir = strRet
        
End Function

Public Function gs_MacroDir(ByVal sMacroName As String, Optional ByVal bIncludeBackslash As Boolean = True, Optional sMacroPath As String = vbNullString) As String

        Dim objVB                   As VBE
        Dim objProject              As VBProject
        Dim strRet                  As String
    
        strRet = vbNullString
        
        Set objVB = App.VBE
        
        For Each objProject In objVB.VBProjects
                With objProject
                        If (StrComp(.name, sMacroName, vbTextCompare) = 0) Then
                                sMacroPath = .FileName
                                strRet = gs_ParseDirName(.FileName, bIncludeBackslash)
                                Exit For
                        End If
                End With
        Next objProject
        
        Set objProject = Nothing
        Set objVB = Nothing
        
        gs_MacroDir = strRet

End Function

Public Function gs_AppDir(Optional ByVal bIncludeBackslash As Boolean = True) As String
        
        Dim strRet                  As String
        
        strRet = App.Path
        
        If bIncludeBackslash Then strRet = gs_EnsureBackslash(strRet)
        
        gs_AppDir = strRet
        
End Function

Public Function gs_ThisDir(Optional ByVal bIncludeBackslash As Boolean = True, Optional sMacroPath As String = vbNullString) As String

        Dim objVB                   As VBE
        Dim objProject              As VBProject
        Dim strRet                  As String
    
        ' the App.Frame.PathOfThisAddin does not always work
        ' if this macro is being utilized while another macro
        ' also being utilized
    
        strRet = vbNullString
                
        Set objVB = App.VBE
        
        For Each objProject In objVB.VBProjects
                With objProject
                        If (StrComp(.name, DEF_MACRO_NAME, vbTextCompare) = 0) Then
                                sMacroPath = .FileName
                                strRet = gs_ParseDirName(.FileName, bIncludeBackslash)
                                Exit For
                        End If
                End With
        Next objProject
        
        Set objProject = Nothing
        Set objVB = Nothing
        
        gs_ThisDir = strRet

End Function

Public Function gs_ThisFile() As String

        Dim objVB                   As VBE
        Dim objProject              As VBProject
        Dim strRet                  As String
    
        strRet = vbNullString
                
        Set objVB = App.VBE
        
        For Each objProject In objVB.VBProjects
                With objProject
                        If (StrComp(.name, DEF_MACRO_NAME, vbTextCompare) = 0) Then
                                strRet = .FileName
                                Exit For
                        End If
                End With
        Next objProject
        
        Set objProject = Nothing
        Set objVB = Nothing
        
        gs_ThisFile = strRet

End Function

Public Function gb_ProjectExists(ByVal sMacroName As String, Optional sReturnFileName As String = vbNullString) As Boolean
    
        Dim vbaProjects             As VBProjects
        Dim vbaProject              As VBProject
        Dim blnRet                  As Boolean
    
On Error Resume Next

        blnRet = False
        
        Set vbaProjects = App.VBE.VBProjects
        
        For Each vbaProject In vbaProjects
                                               
                If (StrComp(sMacroName, vbaProject.name, vbTextCompare) = 0) Then
                        
                        ' will get error if project not saved
                        If (Err.Number = 76) Then
                                Err.Clear
                        Else
                                sReturnFileName = vbaProject.FileName
                                blnRet = True
                                Exit For
                        End If
                    
                End If
    
        Next vbaProject
        
Controlled_Exit:
        
        gb_ProjectExists = blnRet
    
        Set vbaProject = Nothing
        Set vbaProjects = Nothing
    
Exit Function

End Function

Public Function gb_EnsureDirExistance(sRetPath As String) As Boolean
    
        Dim FSO                     As New Scripting.FileSystemObject
        Dim fsoFolder               As Scripting.Folder
        Dim blnRet                  As Boolean
                
On Error GoTo ErrTrap
                
        blnRet = False
                
        ' check if already exists
        If Not FSO.FolderExists(sRetPath) Then
                
                ' doesn't exist so try to create it
                Set fsoFolder = FSO.CreateFolder(sRetPath)
            
                ' if didn't create it then be careful
                If (fsoFolder Is Nothing) Then
                        
                        ' return nothing
                        sRetPath = vbNullString
                        GoTo Controlled_Exit
                        
                End If
                                        
        End If
        
        ' made if here so should be cool
        blnRet = True
        
        ' set return path ensuring we have a backslash at the end of it
        sRetPath = gs_EnsureBackslash(sRetPath)
    
Controlled_Exit:
            
        Set FSO = Nothing
        Set fsoFolder = Nothing
        
        gb_EnsureDirExistance = blnRet
    
Exit Function

ErrTrap:
    
        MsgBox Err.Description, vbExclamation
        blnRet = False
        sRetPath = vbNullString
        Resume Controlled_Exit
    
End Function

Public Function gs_UniqueFileName(ByVal sFile As String) As String
        
        Dim FSO                     As Scripting.FileSystemObject
        Dim lngIndex                As Long
        Dim strExt                  As String
        Dim strFile                 As String
        Dim strFolder               As String
        Dim strTest                 As String
        Dim strRet                  As String
        
        lngIndex = 1
        
        strRet = sFile
        
        strExt = gs_ParseFileExtension(sFile, True)
        strFile = gs_ParseFileName(sFile, True)
        strFolder = gs_ParseDirName(sFile, True)
        
        Set FSO = New Scripting.FileSystemObject
        
        Do While FSO.FileExists(strRet)
                                
                strTest = strFile & " (" & lngIndex & ")" & strExt
                strTest = strFolder & strTest
                
                strRet = strTest
                
                lngIndex = (lngIndex + 1)
        
        Loop
        
        gs_UniqueFileName = strRet
        
        Set FSO = Nothing

End Function

Public Function gl_FileOpenSaveDialogCallbackEx(ByVal lhWnd As LongPtr, ByVal lMsg As Long, _
                                                ByVal lParam As LongPtr, ByVal lpData As LongPtr) As LongPtr
    
        Dim lngHeight               As Long
        Dim lngWidth                As Long
        Dim lngHwnd                 As LongPtr
        Dim lngRet                  As Long
        Dim udtDialog               As RECT
        Dim udtDesktop              As RECT
    
On Error Resume Next
    
        Select Case lMsg
            
                Case WM_INITDIALOG
                
                        ' center the window
                        lngHwnd = GetParent(lhWnd)
                        
                        Call GetWindowRect(GetDesktopWindow, udtDesktop)
                        Call GetWindowRect(lngHwnd, udtDialog)
                        
                        lngHeight = (udtDialog.Bottom - udtDialog.Top)
                        lngWidth = (udtDialog.Right - udtDialog.Left)
                        udtDialog.Left = (((udtDesktop.Right - udtDesktop.Left) - lngWidth) / 2)
                        udtDialog.Top = (((udtDesktop.Bottom - udtDesktop.Top) - lngHeight) / 2)
                        
                        lngRet = MoveWindow(lngHwnd, udtDialog.Left, udtDialog.Top, lngWidth, lngHeight, 1)
                        
        End Select
    
        gl_FileOpenSaveDialogCallbackEx = 0&
        
End Function

Public Function gl_FileOpenSaveDialogCallback(ByVal lhWnd As LongPtr, ByVal lMsg As Long, _
                                              ByVal lParam As LongPtr, ByVal lpData As LongPtr) As LongPtr
    
        Dim lngHeight               As Long
        Dim lngWidth                As Long
        Dim lngHwnd                 As LongPtr
        Dim lngRet                  As Long
        Dim udtDialog               As RECT
        Dim udtDesktop              As RECT
    
On Error Resume Next
    
        Select Case lMsg
            
                Case WM_INITDIALOG
                
                        Call GetWindowRect(GetDesktopWindow, udtDesktop)
                        Call GetWindowRect(lngHwnd, udtDialog)
                        
                        lngHeight = (udtDialog.Bottom - udtDialog.Top)
                        lngWidth = (udtDialog.Right - udtDialog.Left)
                        udtDialog.Left = (((udtDesktop.Right - udtDesktop.Left) - lngWidth) / 2)
                        udtDialog.Top = (((udtDesktop.Bottom - udtDesktop.Top) - lngHeight) / 2)
                        
                        lngRet = MoveWindow(lngHwnd, udtDialog.Left, udtDialog.Top, lngWidth, lngHeight, 1)
                        
        End Select

        gl_FileOpenSaveDialogCallback = 0&
        
End Function

Private Function ml_GetAddressofFunction(add As LongPtr) As LongPtr
        ml_GetAddressofFunction = add
End Function

Private Function ml_BrowseCallbackProc(ByVal lhWnd As LongPtr, ByVal lMsg As Long, ByVal lPIDList As LongPtr, ByVal lData As LongPtr) As LongPtr

        Dim lngPtrRet                     As LongPtr
        Dim lngRet                         As Long
        Dim strBuffer                     As String
    
On Local Error Resume Next

        Select Case lMsg
    
                Case BFFM_INITIALIZED
                
                        lngPtrRet = SendMessage(lhWnd, BFFM_SETSELECTION, 1, m_sCurrentDir)
                        
                Case BFFM_SELCHANGED
                        
                        strBuffer = Space(MAX_PATH)
                        lngRet = SHGetPathFromIDList(lPIDList, strBuffer)

                        If (lngRet = 1) Then
                                lngPtrRet = SendMessage(lhWnd, BFFM_SETSTATUSTEXT, 0, strBuffer)
                        End If

        End Select

        ml_BrowseCallbackProc = 0

End Function

-------------------------------------------------------------------------------
