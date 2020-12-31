' Description: Returns configuration information for all the Terminal Service terminals on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TerminalSetting")

For Each objItem in colItems
  Wscript.Echo "Caption: " & objItem.Caption
  Wscript.Echo "Description: " & objItem.Description
  Wscript.Echo "Setting ID: " & objItem.SettingID
  Wscript.Echo "Terminal name: " & objItem.TerminalName
Next

