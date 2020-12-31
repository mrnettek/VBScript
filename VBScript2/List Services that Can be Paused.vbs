' Description: Returns a list of services that can be paused.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where AcceptPause = True")

For Each objService in colServices
    Wscript.Echo objService.DisplayName 
Next

