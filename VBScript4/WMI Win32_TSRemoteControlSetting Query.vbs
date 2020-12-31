On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSRemoteControlSetting",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "LevelOfControl: " & objItem.LevelOfControl
    Wscript.Echo "RemoteControlPolicy: " & objItem.RemoteControlPolicy
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "TerminalName: " & objItem.TerminalName
Next

