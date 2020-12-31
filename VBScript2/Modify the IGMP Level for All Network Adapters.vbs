' Description: Disables IGMP multicasting on a computer. To enable IP multicasting, pass the value 1 to the SetIGMPLevel method. Pass the value 2 to allow both IP and IGMP multicasting.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetIGMPLevel(0)

