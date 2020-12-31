strComputer = "."

Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
 
objWMIService.Create "Notepad.exe", null, null, intProcessID

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService. _
    ExecNotificationQuery("Select * From __InstanceDeletionEvent " _ 
            & "Within 1 Where TargetInstance ISA 'Win32_Process'")
Do 
    Set objProcess = colItems.NextEvent
    If objProcess.TargetInstance.ProcessID = intProcessID Then
        Exit Do
    End If
Loop

Set objWMIService = GetObject("winmgmts:{(Shutdown)}\\" & _
        strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
 
For Each objItem in colItems
    objItem.Win32Shutdown(0)
Next
  


