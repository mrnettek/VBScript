' Description: Lists the network adapters that can be configured for Terminal Services.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSNetworkAdapterListSetting")

For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Network Adapter ID: " & objItem.NetworkAdapterID
    Wscript.Echo "Network Adapter IP Address: " & objItem.NetworkAdapterIP
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Terminal name: " & objItem.TerminalName
    Wscript.Echo
Next

