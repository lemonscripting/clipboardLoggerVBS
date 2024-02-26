Dim objShell
Set objShell = CreateObject("WScript.Shell")

' Get the current directory
strCurrentDirectory = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")

' Define the relative path to the VBScript file you want to run
strRelativePath = "your_script.vbs"

' Combine the current directory with the relative path
strScriptPath = strCurrentDirectory & "\" & strRelativePath

' Run the VBScript file
objShell.Run strScriptPath

Set objShell = Nothing
