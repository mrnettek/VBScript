On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSNetworkAdapterSetting",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "MaximumConnections: " & objItem.MaximumConnections
    Wscript.Echo "NetworkAdapterID: " & objItem.NetworkAdapterID
    Wscript.Echo "NetworkAdapterName: " & objItem.NetworkAdapterName
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "TerminalName: " & objItem.TerminalName
Next

