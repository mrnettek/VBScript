' Description: Configures the number of forward packets allocated to the router packet queue. Valid values range from 1 to 0xFFFFFFFE.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetNumForwardPackets(1)

