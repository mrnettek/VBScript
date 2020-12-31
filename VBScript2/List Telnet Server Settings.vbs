' Description: Displays Services for UNIX Telnet server settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from TelnetServer_Settings")

For Each objItem in colItems
    Wscript.Echo "Alt Key Mapping: " & objItem.AltKeyMapping
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Default Domain: " & objItem.DefaultDomain
    Wscript.Echo "Idle Session Timeout: " & _
        objItem.IdleSessionTimeout
    Wscript.Echo "Idle Session Timeout Backup: " & _
        objItem.IdleSessionTimeoutBkup
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Kill All: " & objItem.KillAll
    Wscript.Echo "Maximum Connections: " & objItem.MaxConnections
    Wscript.Echo "Maximum Failed Logins: " & _
        objItem.MaxFailedLogins
    Wscript.Echo "Mode Operation: " & objItem.ModeOperation
    Wscript.Echo "Telnet Port: " & objItem.TelnetPort
    Wscript.Echo
Next

