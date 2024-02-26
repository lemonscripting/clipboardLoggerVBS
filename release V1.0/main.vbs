Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strCurrentDir = objFSO.GetAbsolutePathName(".")

strFileName = "sub.vbs"

strSourceFile = strCurrentDir & "\" & strFileName

strStartupFolder = objShell.SpecialFolders("Startup")

strDestFile = strStartupFolder & "\" & strFileName

If objFSO.FileExists(strSourceFile) Then
    objFSO.CopyFile strSourceFile, strDestFile
    WScript.Echo "File copied successfully to Startup folder."
Else
    WScript.Echo "Source file does not exist."
End If
