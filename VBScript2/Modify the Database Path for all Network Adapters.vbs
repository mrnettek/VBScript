' Description: Sets the path to the standard Internet database files (HOSTS, LMHOSTS, NETWORKS, PROTOCOLS) used by the Windows Sockets interface.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetDatabasePath("c:\windows\system32\drivers\etc")

