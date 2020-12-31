strComputer = "."
Set objWMIService = GetObject("winmgmts:{(Security)}\\" & strComputer & "\root\cimv2")

Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'System' " & _
        "AND RecordNumber = 1")

For Each objEvent in colEvents
    Wscript.Echo "Time Written: " & objEvent.TimeWritten
Next
  


