' Description: Sets the TCP window size for all network adapters on a computer. Valid windows sizes range from 0 to 65,535 bytes.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetTCPWindowSize(65535)

