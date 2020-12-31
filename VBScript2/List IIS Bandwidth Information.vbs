' Description: Returns bandwidth information for an IIS server.


strComputer = "LocalHost"
Set objIIS = GetObject("IIS://" & strComputer & "")

Wscript.Echo "Maximum Bandwidth: " & objIIS.MaxBandwidth
Wscript.Echo "Maximum Bandwidth Blocked: " & objIIS.MaxBandwidthBlocked

