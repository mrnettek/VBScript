' Description: Configures the number of times TCP will retransmit a connection attempt before abandoning the effort. Valid values range from 0 to 0xFFFFFFFF.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetTCPMaxConnectRetransmissions(10)

