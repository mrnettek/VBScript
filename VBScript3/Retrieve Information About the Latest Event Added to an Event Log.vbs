strComputer = "."
Set objWMIService = GetObject("winmgmts:{(Security)}\\" & _
        strComputer & "\root\cimv2")

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='System'")

For Each objLogFile in colLogFiles
    intTotal = objLogFile.NumberOfRecords
Next

Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'System' " & _
        "AND RecordNumber = " & intTotal)

For Each objEvent in colEvents
    Wscript.Echo "Time Written: " & objEvent.TimeWritten
Next
  


