' Description: Configures a computer to use zero-broadcasts (0.0.0.0) rather than one-broadcasts (255.255.255.255). Zero-broadcasts are the default used on systems derived from BSD implementations.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetIPUseZeroBroadcast(True)

