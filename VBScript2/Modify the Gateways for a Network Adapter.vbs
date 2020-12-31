' Description: Configures two gateways -- 192.168.1.100 and 192.168.1.200 -- for a network adapter. Note that even if only one gateway is specified, the IP address for that gateway must be passed as part of an array (in that case, an array with only a single element).


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colNetCards = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objNetCard in colNetCards
    arrGateways = Array("192.168.1.100", "192.168.1.200")
    objNetCard.SetGateways(arrGateways)
Next

