' Description: Sets the default time-to-live value in the header of outgoing IP packets to 64 (this represents the number of routers an IP packet can pass through before being discarded). Valid TTL values range from 1 to 255, with a default value of 32.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetDefaultTTL(64)

