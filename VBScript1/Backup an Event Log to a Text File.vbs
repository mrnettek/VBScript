strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where LogFile='Application'")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("C:\Scripts\Events.txt")

For Each objEvent in colEvents
    strTimeWritten = objEvent.TimeWritten

    dtmTimeWritten = CDate(Mid(strTimeWritten, 5, 2) & "/" & _
        Mid(strTimeWritten, 7, 2) & "/" & Left(strTimeWritten, 4) _
            & " " & Mid (strTimeWritten, 9, 2) & ":" & _
                Mid(strTimeWritten, 11, 2) & ":" & Mid(strTimeWritten, 13, 2))

    dtmDate = FormatDateTime(dtmTimeWritten, vbShortDate)
    dtmTime = FormatDateTime(dtmTimeWritten, vbLongTime)

    strEvent = dtmDate & vbTab
    strEvent = strEvent & dtmTime & vbTab
    strEvent = strEvent & objEvent.SourceName & vbTab
    strEvent = strEvent & objEvent.Type & vbTab
    strEvent = strEvent & objEvent.Category & vbTab
    strEvent = strEvent & objEvent.EventCode & vbTab
    strEvent = strEvent & objEvent.User & vbTab
    strEvent = strEvent & objEvent.ComputerName & vbTab

    strDescription = objEvent.Message
    If IsNull(strDescription) Then
        strDescription = "The event description cannot be found."
    End If
    strDescription = Replace(strDescription, vbCrLf, " ")
    strEvent = strEvent & strDescription

    objFile.WriteLine strEvent
Next

objFile.Close
  


