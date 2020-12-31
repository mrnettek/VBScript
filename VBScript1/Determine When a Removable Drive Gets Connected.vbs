strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colEvents = objWMIService.ExecNotificationQuery _
    ("Select * From __InstanceOperationEvent Within 10 Where " _
        & "TargetInstance isa 'Win32_LogicalDisk'")

Do While True
    Set objEvent = colEvents.NextEvent
    If objEvent.TargetInstance.DriveType = 2 Then 
        Select Case objEvent.Path_.Class
            Case "__InstanceCreationEvent"
                Wscript.Echo "Drive " & objEvent.TargetInstance.DeviceId & _
                    " has been added."
            Case "__InstanceDeletionEvent"
                Wscript.Echo "Drive " & objEvent.TargetInstance.DeviceId & _
                    " has been removed."
        End Select
    End If
 Loop
  


