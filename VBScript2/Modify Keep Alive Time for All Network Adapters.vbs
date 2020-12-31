' Description: Configures the keep alive time for all network adapters on a computer to 300,000 milliseconds (5 minutes).


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetKeepAliveTime(300000)

