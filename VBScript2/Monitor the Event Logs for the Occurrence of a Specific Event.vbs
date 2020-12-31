strComputer = "."

Set objWMIService = GetObject("winmgmts:{(Security)}\\" & _
        strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _    
    ("Select * from __InstanceCreationEvent Where " _
        & "TargetInstance ISA 'Win32_NTLogEvent' " _
            & "and TargetInstance.EventCode = '0' ")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
    Wscript.Echo objLatestEvent.TargetInstance.User
    Wscript.Echo objLatestEvent.TargetInstance.TimeWritten
    Wscript.Echo objLatestEvent.TargetInstance.Message
    Wscript.Echo
Loop
  


