arrProcesses = Array("freecell.exe","sol.exe","spider.exe","winmine.exe")

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

i = 0

Set colMonitoredProcesses = objWMIService. ExecNotificationQuery _        
    ("Select * From __InstanceCreationEvent Within 5 Where TargetInstance ISA 'Win32_Process'")

Do While i = 0
    Set objLatestProcess = colMonitoredProcesses.NextEvent
    strProcess = LCase(objLatestProcess.TargetInstance.Name)

    For Each strName in arrProcesses
        If strName = strProcess Then
            Wscript.Echo strName & " has started."
        End If
    Next
Loop
  


