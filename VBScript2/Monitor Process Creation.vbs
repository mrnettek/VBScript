' Description: Temporary event consumer that issues an alert each time a new process is created.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colMonitoredProcesses = objWMIService. _        
    ExecNotificationQuery("select * from __instancecreationevent " _ 
        & " within 1 where TargetInstance isa 'Win32_Process'")
i = 0

Do While i = 0
    Set objLatestProcess = colMonitoredProcesses.NextEvent
    Wscript.Echo objLatestProcess.TargetInstance.Name
Loop

