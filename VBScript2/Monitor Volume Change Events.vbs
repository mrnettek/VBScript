' Description: Temporary event consumer that issues an alert when local volumes are added to or deleted from a computer. (This class monitors only changes to local drives; it cannot detect the addition/deletion of network volumes.)


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService. _
    ExecNotificationQuery("Select * from Win32_VolumeChangeEvent")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
    Wscript.Echo objLatestEvent.DriveName
    Wscript.Echo objLatestEvent.EventType
    Wscript.Echo objLatestEvent.Time_Created
Loop

