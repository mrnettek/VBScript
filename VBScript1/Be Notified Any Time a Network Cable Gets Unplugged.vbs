strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\wmi")
Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("Select * from MSNdis_StatusMediaDisconnect") 

Do While True 
    Set strLatestEvent = colMonitoredEvents.NextEvent 
    Wscript.Echo "A network connection has been lost:"
    WScript.Echo strLatestEvent.InstanceName, Now
    Wscript.Echo 
 Loop
  


