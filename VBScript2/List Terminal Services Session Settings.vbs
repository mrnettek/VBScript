' Description: Returns Terminal Service session configuration information.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSSessionSetting")

For Each objItem in colItems
    Wscript.Echo "Active session limit: " & objItem.ActiveSessionLimit
    Wscript.Echo "Broken connection action: " & objItem.BrokenConnectionAction
    Wscript.Echo "Broken connection policy: " & objItem.BrokenConnectionPolicy
    Wscript.Echo "Disconnected session limit: " & _
        objItem.DisconnectedSessionLimit
    Wscript.Echo "Idle session limit: " & objItem.IdleSessionLimit
    Wscript.Echo "Reconnection policy: " & objItem.ReconnectionPolicy
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Terminal name: " & objItem.TerminalName
    Wscript.Echo "Time limit policy: " & objItem.TimeLimitPolicy
    Wscript.Echo
Next

