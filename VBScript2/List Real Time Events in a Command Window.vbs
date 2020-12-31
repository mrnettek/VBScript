' Description: Creates a temporary event consumer that monitors the event log for error events. When an error event occurs, the script displays the event information in the command window.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Security)}!\\" & _
        strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("Select * from __InstanceCreationEvent within 5 where TargetInstance " & _
        isa 'Win32_NTLogEvent' and TargetInstance.EventType = '1'")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
        Wscript.Echo "Record No.: " & _
            objLatestEvent.TargetInstance.RecordNumber
        Wscript.Echo "Event ID: " & objLatestEvent.TargetInstance.EventCode
        Wscript.Echo "Time: " & objLatestEvent.TargetInstance.TimeWritten
        Wscript.Echo "Source: " & objLatestEvent.TargetInstance.SourceName
        Wscript.Echo "Category: " & _
            objLatestEvent.TargetInstance.CategoryString
        Wscript.Echo "Event Type: " & objLatestEvent.TargetInstance.Type
        Wscript.Echo "Computer: " & _
            objLatestEvent.TargetInstance.ComputerName
        Wscript.Echo "User: " & objLatestEvent.TargetInstance.User
        Wscript.echo "Text: " & objLatestEvent.TargetInstance.Message
Loop

