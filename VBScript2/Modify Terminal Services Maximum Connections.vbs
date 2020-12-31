' Description: Sets the maximum number of connections allowed on a Terminal Services network adapter to 100.


Const MAXIMUM_CONNECTIONS = 100
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TSNetworkAdapterSetting")

For Each objItem in colItems
    objItem.MaximumConnections = MAXIMUM_CONNECTIONS
    objItem.Put_
Next

