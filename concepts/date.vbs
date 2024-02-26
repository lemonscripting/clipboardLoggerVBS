Function SERVER_ACCESS_STATUS()
    Dim objFSO, objFile, strFilePath, strToday, strFileContent

    strScriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
    strToday = Right("0" & Day(Date), 2) & "/" & Right("0" & Month(Date), 2) & "/" & Year(Date)
    strFilePath = strScriptDir & "send_dates.txt"

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

    Wscript.Echo SERVER_ACCESS_STATUS()

'Case 1
'data already sent to server today
'Case 2 
'data not sent to server today
'Case 3 
'initial run of the script, data not sent to server today

Select Case SERVER_ACCESS_STATUS()
    Case 1
        ' Code block to execute if SERVER_ACCESS_STATUS() is 1
        Wscript.Echo SERVER_ACCESS_STATUS()
    Case 2
        ' Code block to execute if SERVER_ACCESS_STATUS() is 2
        Wscript.Echo SERVER_ACCESS_STATUS()
        'SEND_MAIL()
    Case 3
        ' Code block to execute if SERVER_ACCESS_STATUS() is 3
        Wscript.Echo SERVER_ACCESS_STATUS()
        'SEND_MAIL()
    Case Else
        ' Code block to execute if SERVER_ACCESS_STATUS() does not match any of the specified cases
        Wscript.Echo SERVER_ACCESS_STATUS()
End Select
