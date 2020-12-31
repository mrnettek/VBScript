' Description: Temporary event consumer that issues an alert any time a service changes status (for example, an active service that is paused or stopped).


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServices = objWMIService. _ 
    ExecNotificationQuery("Select * from __instancemodificationevent " _ 
        & "within 30 where TargetInstance isa 'Win32_Service'")
i = 0

Do While i = 0
    Set objService = colServices.NextEvent
    If objService.TargetInstance.State <> _ 
        objService.PreviousInstance.State Then
        Wscript.Echo objService.TargetInstance.Name _ 
            &  " is " & objService.TargetInstance.State _
                & ". The service previously was " & _
                    objService.PreviousInstance.State & "."
    End If
Loop

