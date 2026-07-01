Attribute VB_Name = "Events"
Option Explicit
'Option Private Module
'

Function InitAlphacamAddIn(AcamVersion As Long) As Integer
                
        Dim FSO                     As New Scripting.FileSystemObject
        Dim sNestTitle              As String
        Dim sBMP                    As String
        Dim ID                      As Integer
On Error Resume Next
        
        ' 08 jun 10 - rg
        '   + MODIFIED to below
        '
        'fr.AddMenuItem2 App.Frame.ReadTextFile("ReverseNest.txt", 1, 1), "RevSide", acamMenuSPECIALFUN, ""
        
        If (gi_NestLevel <> 1) Then Exit Function
        
        With App.Frame
                ID = .CreateButtonBar("REVNST")
                .AddMenuItem2 "&参数设置", "FuncSettings", acamMenuNEW, "反面加工"
                .AddMenuItem2 "&输出路径", "OutPutNC", acamMenuNEW, "反面加工"
                
                sNestTitle = "REVNST"

'                If .AddMenuItem3(gv_CTX(1, 1, "Reverse-Side Nesting"), "g_RevSide", acamMenuUTILS_NEST, "", sNestTitle) Then
'                        sBMP = gs_ThisDir & "ReverseNest.bmp"
'                        If FSO.FileExists(sBMP) Then Call .AddButton(acamButtonBarUTILS, sBMP, .LastMenuCommandID)
'                End If

        End With
            
        InitAlphacamAddIn = 0
        
End Function
Function FuncSettings()
    Shell App.Path & "\settings.exe", vbNormalFocus
End Function

Function OutPutNC()
    Shell App.Path & "\output.exe", vbNormalFocus
End Function
Public Function g_RevSide()
        
On Error Resume Next
        App.ActiveDrawing.SetGeosSelected False
        App.ActiveDrawing.SetToolPathsSelected False
        
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
