' Description: Configures the number of times TCP will attempt to retransmit an individual data segment before abandoning the effort. Valid values range from 0 to 0xFFFFFFFF.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetTCPMaxDataRetransmissions(10)

