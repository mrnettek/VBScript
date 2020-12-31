' Description: Displays Services for UNIX Telnet server auditing settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Telnet_Auditing")

For Each objItem in colItems
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Event Logging Enabled: " & _
        objItem.EventLoggingEnabled
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Log Administrator Attempts: " & _
        objItem.LogAdminAttempts
    Wscript.Echo "Log Events: " & objItem.LogEvents
    Wscript.Echo "Log Failures: " & objItem.LogFailures
    Wscript.Echo "Log File: " & objItem.LogFile
    Wscript.Echo "Log File Size: " & objItem.LogFileSize
    Wscript.Echo "Log To File: " & objItem.LogToFile
    Wscript.Echo
Next

