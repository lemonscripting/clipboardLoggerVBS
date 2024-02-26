Const ForAppending = 8
Const DelaySeconds = 1

Set objShell = CreateObject("WScript.Shell")
strFolder = objShell.CurrentDirectory
strFilePath = strFolder & "\cliplog.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Function to get current date and time in the specified format
Function GetCurrentDateTime()
    Dim dt
    dt = Now
    GetCurrentDateTime = FormatDateTime(dt, vbShortDate) & " " & FormatDateTime(dt, vbShortTime)
End Function

' Function to get the newest clipboard content
Function GetNewestClipboardContent()
    Set objHTML = CreateObject("htmlfile")
    Set objData = objHTML.ParentWindow.ClipboardData

    GetNewestClipboardContent = objData.GetData("text")
End Function

' Main procedure to monitor clipboard and update log file
Sub MonitorClipboard()
    Dim strDateTime, strContent, strPrevContent

    ' Initialize previous clipboard content
    strPrevContent = ""

    ' Infinite loop to continuously monitor clipboard
    Do
        ' Get current clipboard content
        strContent = GetNewestClipboardContent()

        ' Check if clipboard content has changed
        If strContent <> strPrevContent Then
            ' Get current date and time
            strDateTime = "[" & GetCurrentDateTime() & "]"

            ' Create or open the log file
            Set objLogFile = objFSO.OpenTextFile(strFilePath, ForAppending, True)

            ' Write the date, time, and clipboard content to the log file
            objLogFile.WriteLine strDateTime & " " & strContent

            ' Close the log file
            objLogFile.Close

            ' Update previous clipboard content
            strPrevContent = strContent
        End If

        ' Wait for a short delay before checking clipboard again
        WScript.Sleep DelaySeconds * 1000
    Loop
End Sub

' Call the main procedure to start monitoring clipboard
MonitorClipboard
