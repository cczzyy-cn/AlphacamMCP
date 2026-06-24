in file: C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\000\StartUp\Utils\ReverseNest\ReverseNest.amb - OLE stream: 'vao/The VBA Project/_VBA_Project/VBA/Events'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Option Private Module
'

Function InitAlphacamAddIn(AcamVersion As Long) As Integer
                
        Dim FSO                     As New Scripting.FileSystemObject
        Dim sNestTitle              As String
        Dim sBMP                    As String
        
On Error Resume Next
        
        ' 08 jun 10 - rg
        '   + MODIFIED to below
        '
        'fr.AddMenuItem2 App.Frame.ReadTextFile("ReverseNest.txt", 1, 1), "RevSide", acamMenuSPECIALFUN, ""
        
        If (gi_NestLevel <> 1) Then Exit Function
        
        With App.Frame
        
                sNestTitle = gs_ReadAcamNestCTX(1140, 1, "&Nesting")
                
                If .AddMenuItem3(gv_CTX(1, 1, "Reverse-Side Nesting"), "g_RevSide", acamMenuUTILS_NEST, "", sNestTitle) Then
                        sBMP = gs_ThisDir & "ReverseNest.bmp"
                        If FSO.FileExists(sBMP) Then Call .AddButton(acamButtonBarUTILS, sBMP, .LastMenuCommandID)
                End If
        
        End With
            
        InitAlphacamAddIn = 0
        
End Function

Public Function g_RevSide()
        
On Error Resume Next
        
        Call Load(frmMain)
        DoEvents
        Call frmMain.Show
        
End Function

Public Function OnUpdateg_RevSide()
        
        Dim intRet                  As Integer
        
On Error GoTo ErrTrap
        
        intRet = IIf(gb_HasNesting(Nothing), 1, 0)

Controlled_Exit:

        OnUpdateg_RevSide = intRet

Exit Function
        
ErrTrap:
        
        intRet = 0
        Resume Controlled_Exit

End Function
-------------------------------------------------------------------------------
