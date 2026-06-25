in file: C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\000\StartUp\Utils\ReverseNest\ReverseNest.amb - OLE stream: 'vao/The VBA Project/_VBA_Project/VBA/frmMain'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 




Option Explicit
'

Private Sub cmdCancel_Click()
        
On Error Resume Next
        
        Call Me.Hide
        DoEvents
        Call Unload(Me)

End Sub

Private Sub cmdOK_Click()
        
        Dim blnRet                  As Boolean
        
        Call Me.Hide
        DoEvents
        
        If optFlipX.Value Then
                Call g_AroundX(optSheetOrderBySheet.Value, chkMinChanges.Value)
        Else
                Call g_AroundY(optSheetOrderBySheet.Value, chkMinChanges.Value)
        End If
        
        ' save last settings
        blnRet = gb_WriteRegKey(REG_SZ, DEF_REG_KEY, "OrderBySheet", Abs(optSheetOrderBySheet.Value))
        blnRet = gb_WriteRegKey(REG_SZ, DEF_REG_KEY, "FlipAroundX", Abs(optFlipX.Value))
        blnRet = gb_WriteRegKey(REG_SZ, DEF_REG_KEY, "MinToolChanges", Abs(chkMinChanges.Value))
        
        '01 OCT 10 +SDO
        blnRet = gb_WriteRegKey(REG_SZ, DEF_REG_KEY, "IncludeGeos", Abs(chkGeos.Value))
        
        Call Unload(Me)

End Sub

Private Sub UserForm_Initialize()
        
        Dim I As Integer
                
On Error Resume Next
                
        Me.Caption = gv_CTX(10, 1, Me.Caption)
        
        fraSheetOrder.Caption = gv_CTX(10, 2, fraSheetOrder.Caption)
        optSheetOrderBySide.Caption = gv_CTX(20, 1, optSheetOrderBySide.Caption)
        optSheetOrderBySheet.Caption = gv_CTX(20, 2, optSheetOrderBySheet.Caption)
        
        fraFlip.Caption = gv_CTX(10, 3, fraFlip.Caption)
        optFlipX.Caption = gv_CTX(20, 4, optFlipX.Caption)
        optFlipY.Caption = gv_CTX(20, 5, optFlipY.Caption)
        
        chkMinChanges.Caption = gv_CTX(20, 3, chkMinChanges.Caption)
        
        '01 OCT 10 +SDO
        chkGeos.Caption = gv_CTX(20, 6, chkGeos.Caption)
        
        cmdOK.Caption = gs_ReadAeditCTX(65, 1, cmdOK.Caption)
        cmdCancel.Caption = gs_ReadAeditCTX(65, 2, cmdCancel.Caption)
                    
        ' retrieve last settings
        optSheetOrderBySheet.Value = CBool(gs_ReadRegKey(DEF_REG_KEY, "OrderBySheet", HKEY_CURRENT_USER, GetSetting("LicomSystems", "_ReverseNest", "sheetorder", 0)))
        optSheetOrderBySide.Value = Not optSheetOrderBySheet.Value
        
        optFlipX.Value = CBool(gs_ReadRegKey(DEF_REG_KEY, "FlipAroundX", HKEY_CURRENT_USER, "0"))
        optFlipY.Value = Not optFlipX.Value
        
        chkMinChanges.Value = CBool(gs_ReadRegKey(DEF_REG_KEY, "MinToolChanges", HKEY_CURRENT_USER, GetSetting("LicomSystems", "_ReverseNest", "toolopt", 1)))
        
        '01 OCT 10 +SDO
        chkGeos.Value = CBool(gs_ReadRegKey(DEF_REG_KEY, "IncludeGeos", HKEY_CURRENT_USER, GetSetting("LicomSystems", "_ReverseNest", "IncGeos", 1)))
            
        Call g_SetAccelerators(Me)
        
End Sub

-------------------------------------------------------------------------------
