' Description: Lists Terminal Services network adapter settings, including the maximum number of connections allowed.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSNetworkAdapterSetting")

For Each objItem in colItems
    Wscript.Echo "Maximum Connections: " & objItem.MaximumConnections
    Wscript.Echo "Network Adapter ID: " & objItem.NetworkAdapterID
    Wscript.Echo "Network Adapter Name: " & objItem.NetworkAdapterName
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Terminal Name: " & objItem.TerminalName
    Wscript.Echo
Next

