' Description: Temporary event consumer that issues an alert any time a file is created in the C:\Scripts folder. Best when run under Cscript.exe.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("SELECT * FROM __InstanceCreationEvent WITHIN 10 WHERE " _
        & "Targetinstance ISA 'CIM_DirectoryContainsFile' and " _
            & "TargetInstance.GroupComponent= " _
                & "'Win32_Directory.Name=""c:\\\\scripts""'")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
    Wscript.Echo objLatestEvent.TargetInstance.PartComponent
Loop

