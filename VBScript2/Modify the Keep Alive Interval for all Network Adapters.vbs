' Description: Configures the keep alive interval for all network adapters on a computer to 300,00 milliseconds (5 minutes).


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetKeepAliveInterval(300000)

