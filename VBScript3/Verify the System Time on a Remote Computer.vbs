strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * From Win32_LocalTime")
 
For Each objItem in colItems
    strTime = objItem.Hour & ":" & objItem.Minute & ":" & objItem.Second
    dtmTime = CDate(strTime)
    Wscript.Echo FormatDateTime(dtmTime, vbFormatLongTime)
Next


