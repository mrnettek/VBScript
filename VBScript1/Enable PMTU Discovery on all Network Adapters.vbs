' Description: Enables a computer to automatically discover the maximum transmission unit on a network. To disable auto-discovery of the maximum transmission unit, pass the value False to the SetPMTUDiscovery method.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetPMTUDiscovery(True)

