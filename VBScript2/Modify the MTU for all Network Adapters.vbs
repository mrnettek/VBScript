' Description: Configures the maximum transmission unit for all network adapters installed in a computer. Valid values range from a minimum of 68 bytes to the maximum number of bytes supported by the network.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetMTU(68)

