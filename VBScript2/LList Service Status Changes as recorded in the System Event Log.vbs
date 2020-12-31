' Description: Retrieves events from the System event log that have an event ID of 7036. These events are recorded any time a service changes status.


Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServiceEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'System' and " _
        & "EventCode = '7036'")

For Each strEvent in colServiceEvents
    dtmConvertedDate.Value = strEvent.TimeWritten
    Wscript.Echo dtmConvertedDate.GetVarDate    
    Wscript.Echo strEvent.Message
Next

