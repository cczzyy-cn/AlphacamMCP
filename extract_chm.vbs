' 使用 Windows HTML Help API 提取 CHM 内容
Dim fso, shell, chmPath, outDir, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

chmPath = "C:\Program Files (x86)\Vero Software\Alphacam 2016 R1\Add-Ins\OpListJP\OpListJP.chm"
outDir = "C:\Users\C\Desktop\OpListJP_Help"

If Not fso.FolderExists(outDir) Then
    fso.CreateFolder(outDir)
End If

' 使用 hh.exe 反编译 CHM
cmd = "hh.exe -decompile """ & outDir & """ """ & chmPath & """"
shell.Run cmd, 0, True

' 列出提取的文件
Dim folder, file, subFolder
Set folder = fso.GetFolder(outDir)
For Each subFolder In folder.SubFolders
    WScript.Echo "Folder: " & subFolder.Name
    For Each file In subFolder.Files
        WScript.Echo "  " & file.Name & " (" & file.Size & " bytes)"
    Next
Next
' Also list root files
For Each file In folder.Files
    WScript.Echo file.Name & " (" & file.Size & " bytes)"
Next
