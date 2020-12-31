strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objEventSource = objWMIService.ExecNotificationQuery _
    ("SELECT * FROM __InstanceOperationEvent WITHIN 10 WHERE TargetInstance ISA 'Win32_Process'")

Do While True
    Set objEventObject = objEventSource.NextEvent()
    If Right(objEventObject.TargetInstance.Name, 4) = ".scr" Then
        Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
        For Each objItem in colItems
            objItem.Win32Shutdown(4)
        Next
    End If
Loop
  


