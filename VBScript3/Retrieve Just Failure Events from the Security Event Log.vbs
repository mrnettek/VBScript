strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Security)}!\\" & _
        strComputer & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * FROM Win32_NTLogEvent WHERE Logfile = 'Security' " & _
        "AND EventType = 5")
For Each objEvent in colLoggedEvents
    Wscript.Echo "==================================================="
    Wscript.Echo "Category: " & objEvent.Category
    Wscript.Echo "Computer Name: " & objEvent.ComputerName
    Wscript.Echo "Event Code: " & objEvent.EventCode
    Wscript.Echo "Message: " & objEvent.Message
    Wscript.Echo "Record Number: " & objEvent.RecordNumber
    Wscript.Echo "Source Name: " & objEvent.SourceName
    Wscript.Echo "Time Written: " & objEvent.TimeWritten
    Wscript.Echo "Event Type: " & objEvent.Type
    Wscript.Echo "User: " & objEvent.User
    Wscript.Echo
Next
  


