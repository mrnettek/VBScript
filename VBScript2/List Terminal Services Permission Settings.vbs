' Description: Lists Terminal Services permission settings.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSPermissionsSetting")

For Each objItem in colItems
  Wscript.Echo "Caption: " & objItem.Caption
  Wscript.Echo "Description: " & objItem.Description
  Wscript.Echo "Setting ID: " & objItem.SettingID
  Wscript.Echo "Terminal name: " & objItem.TerminalName
Next

