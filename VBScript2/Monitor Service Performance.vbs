' Description: Uses formatted performance counters to retrieve performance data for the DHCP Server service.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colDHCPServer = objRefresher.AddEnum(objWMIService, _
    "Win32_PerfFormattedData_DHCPServer_DHCPServer").ObjectSet
objRefresher.Refresh

For i = 1 to 60
    For Each objDHCPServer in colDHCPServer
        Wscript.Echo "Acknowledgements per second: " & _
            objDHCPServer.AcksPerSec
        Wscript.Echo "Declines per second: " & _
            objDHCPServer.DeclinesPerSec
        Wscript.Echo "Discovers per second: " & _
            objDHCPServer.DiscoversPerSec
        Wscript.Echo "Informs per second: " & objDHCPServer.InformsPerSec
        Wscript.Echo "Offers per second: " & objDHCPServer.OffersPerSec
        Wscript.Echo "Releases per second: " & _
            objDHCPServer.ReleasesPerSec
        Wscript.Echo "Requests per second: " & _
            objDHCPServer.RequestsPerSec
    Next
    Wscript.Sleep 10000
    objRefresher.Refresh
Next

