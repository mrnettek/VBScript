' Description: Issues an alert if a computer changes power state (for example, enters or leaves suspend mode).


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("Select * from Win32_PowerManagementEvent")

Do
    Set strLatestEvent = colMonitoredEvents.NextEvent
    Wscript.Echo strLatestEvent.EventType
Loop

