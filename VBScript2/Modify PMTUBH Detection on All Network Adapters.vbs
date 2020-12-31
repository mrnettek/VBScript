' Description: Enables the auto-discovery of black hole routers when determining the maximum transmission unit on a network. To disable auto-discovery of black hole routers, pass the value False to the SetPMTUBHDetect method.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetPMTUBHDetect(True)

