' Description: Temporary event consumer that issues an alert any time a file is deleted from the C:\Scripts folder. Best when run under Cscript.exe.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("SELECT * FROM __InstanceDeletionEvent WITHIN 10 WHERE " _
        & "Targetinstance ISA 'CIM_DirectoryContainsFile' and " _
            & "TargetInstance.GroupComponent= " _
                & "'Win32_Directory.Name=""c:\\\\scripts""'")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
    Wscript.Echo objLatestEvent.TargetInstance.PartComponent
Loop

