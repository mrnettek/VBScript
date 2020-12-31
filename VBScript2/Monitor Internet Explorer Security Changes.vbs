' Description: Issues an alert any time an audited Internet Explorer setting is changed and recorded in the Security event log.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{"{impersonationLevel=impersonate,(Security)}!\\" & strComputer & _
        "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _    
    ("SELECT * FROM __InstanceCreationEvent WHERE TargetInstance ISA " _
        & "'Win32_NTLogEvent' AND TargetInstance.EventCode = '560' AND " _
            & "TargetInstance.Logfile = 'Security' GROUP WITHIN 2")
Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
        strAlertToSend = "Internet Explorer security settings have been " & _
            "changed."
        Wscript.Echo strAlertToSend
Loop

