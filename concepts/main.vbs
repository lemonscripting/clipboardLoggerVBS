Const ForAppending = 8
Const DelaySeconds = 1

Set objShell = CreateObject("WScript.Shell")
strFolder = objShell.CurrentDirectory
strFilePath = strFolder & "\cliplog.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")

Function GetCurrentDateTime()
    Dim dt
    dt = Now
    GetCurrentDateTime = FormatDateTime(dt, vbShortDate) & " " & FormatDateTime(dt, vbShortTime)
End Function

Function GetNewestClipboardContent()
    Set objHTML = CreateObject("htmlfile")
    Set objData = objHTML.ParentWindow.ClipboardData
    GetNewestClipboardContent = objData.GetData("text")
End Function

Sub MonitorClipboard()
    Dim strDateTime, strContent, strPrevContent
    strPrevContent = ""

    Do
        strContent = GetNewestClipboardContent()

        If strContent <> strPrevContent Then
            strDateTime = "[" & GetCurrentDateTime() & "]"
            Set objLogFile = objFSO.OpenTextFile(strFilePath, ForAppending, True)
            objLogFile.WriteLine strDateTime & vbNewLine & strContent
            objLogFile.Close
            strPrevContent = strContent
        End If

        WScript.Sleep DelaySeconds * 1000
    Loop
End Sub

MonitorClipboard
