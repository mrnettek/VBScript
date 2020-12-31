' Description: Temporary event consumer that issues an alert any time the file C:\Scripts\Index.vbs is modified. Best when run under Cscript.exe.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("SELECT * FROM __InstanceModificationEvent WITHIN 10 WHERE " _
        & "TargetInstance ISA 'CIM_DataFile' and " _
            & "TargetInstance.Name='c:\\scripts\\index.vbs'")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
    Wscript.Echo "File: " & objLatestEvent.TargetInstance.Name
    Wscript.Echo "New size: " & objLatestEvent.TargetInstance.FileSize
    Wscript.Echo "Old size: " & objLatestEvent.PreviousInstance.FileSize
Loop

