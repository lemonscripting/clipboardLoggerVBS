Dim objFSO, objFile, objFolder, objTextFile, objShell

' Create FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Create Shell object to get desktop folder path
Set objShell = CreateObject("WScript.Shell")

' Get desktop folder path
strDesktopPath = objShell.SpecialFolders("Desktop")

' Create output text file on desktop
Set objTextFile = objFSO.CreateTextFile(strDesktopPath & "\FileList.txt")

' Start from the root folder (change as needed)
Set objFolder = objFSO.GetFolder("C:\")

' Recursively traverse through all folders and files
TraverseFolder objFolder

' Close the text file
objTextFile.Close
Set objTextFile = Nothing

' Cleanup
Set objFolder = Nothing
Set objFSO = Nothing

WScript.Echo "File paths have been saved to FileList.txt on your desktop."

Sub TraverseFolder(folder)
    Dim subfolder, file

    ' Process files in current folder
    For Each file In folder.Files
        objTextFile.WriteLine file.Path
    Next

    ' Recursively process subfolders
    For Each subfolder In folder.SubFolders
        TraverseFolder subfolder
    Next
End Sub
