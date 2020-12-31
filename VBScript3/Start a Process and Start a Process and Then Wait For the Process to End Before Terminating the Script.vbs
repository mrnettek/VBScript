strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
objWMIService.Create "notepad.exe", null, null, intProcessID

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colMonitoredProcesses = objWMIService.ExecNotificationQuery _
    ("Select * From __InstanceDeletionEvent Within 1 Where TargetInstance ISA 'Win32_Process'")

Do Until i = 1
    Set objLatestProcess = colMonitoredProcesses.NextEvent
    If objLatestProcess.TargetInstance.ProcessID = intProcessID Then
        i = 1
    End If
Loop

Wscript.Echo "Notepad has been terminated."
  


