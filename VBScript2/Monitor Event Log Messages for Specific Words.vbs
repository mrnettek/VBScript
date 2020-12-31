strComputer = "."

Set objWMIService = GetObject("winmgmts:{(Security)}\\" & _
        strComputer & "\root\cimv2")

Set colEvents = objWMIService.ExecNotificationQuery _    
    ("Select * From __InstanceCreationEvent Where " _
        & "TargetInstance isa 'Win32_NTLogEvent'")

Do
    Set objEvent = colEvents.NextEvent
    If InStr(LCase(objEvent.TargetInstance.Message), "test.vbs") Then
        Wscript.Echo Now
        Wscript.Echo "Category: " & objEvent.TargetInstance.Category
        Wscript.Echo "Event Code: " & objEvent.TargetInstance.EventCode
        Wscript.Echo "Message: " & objEvent.TargetInstance.Message
        Wscript.Echo "Record Number: " & objEvent.TargetInstance.RecordNumber
        Wscript.Echo "Source Name: " & objEvent.TargetInstance.SourceName
        Wscript.Echo "Event Type: " & objEvent.TargetInstance.Type
        Wscript.Echo
    End If
Loop
  


