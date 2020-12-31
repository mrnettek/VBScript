' Description: Returns information about Terminal Service remote control as configured on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSRemoteControlSetting")

For Each objItem in colItems
    Wscript.Echo "Level of control: " & objItem.LevelofControl
    Wscript.Echo "Remote control policy: " & objItem.RemoteControlPolicy
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Terminal name: " & objItem.TerminalName
    Wscript.Echo
Next

