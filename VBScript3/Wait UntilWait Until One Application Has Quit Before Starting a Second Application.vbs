strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objProcess = GetObject("winmgmts:root\cimv2:Win32_Process")

errReturn = objProcess.Create("Notepad.exe", null, null, intProcessID)

Set colMonitoredProcesses = objWMIService. _        
    ExecNotificationQuery("select * From __InstanceDeletionEvent " _ 
        & " within 1 where TargetInstance isa 'Win32_Process'")

Do While True
    Set objLatestProcess = colMonitoredProcesses.NextEvent
    If objLatestProcess.TargetInstance.ProcessID = intProcessID Then
        Exit Do
    End If
Loop

errReturn = objProcess.Create("Calc.exe", null, null, intProcessID)
  


