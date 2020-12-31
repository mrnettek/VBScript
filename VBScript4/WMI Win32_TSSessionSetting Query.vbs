On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSSessionSetting",,48)
For Each objItem in colItems
    Wscript.Echo "ActiveSessionLimit: " & objItem.ActiveSessionLimit
    Wscript.Echo "BrokenConnectionAction: " & objItem.BrokenConnectionAction
    Wscript.Echo "BrokenConnectionPolicy: " & objItem.BrokenConnectionPolicy
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DisconnectedSessionLimit: " & objItem.DisconnectedSessionLimit
    Wscript.Echo "IdleSessionLimit: " & objItem.IdleSessionLimit
    Wscript.Echo "ReconnectionPolicy: " & objItem.ReconnectionPolicy
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "TerminalName: " & objItem.TerminalName
    Wscript.Echo "TimeLimitPolicy: " & objItem.TimeLimitPolicy
Next

