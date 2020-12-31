strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
 
errResult = objWMIService.Create("calc.exe", null, null, intCalcID)
errResult = objWMIService.Create("notepad.exe", null, null, intNotepadID)

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colProcesses = objWMIService.ExecNotificationQuery _
    ("Select * From __InstanceDeletionEvent " _ 
            & "Within 1 Where TargetInstance ISA 'Win32_Process'")

Do Until i = 999
    Set objProcess = colProcesses.NextEvent
    If objProcess.TargetInstance.ProcessID = intCalcID Then
        Exit Do
    End If
Loop

Set colProcesses = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where ProcessID = " & intNotepadID)

For Each objProcess in colProcesses
    objProcess.Terminate()
Next
  


