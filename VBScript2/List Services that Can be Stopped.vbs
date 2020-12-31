' Description: Returns a list of services that can be stopped.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where AcceptStop = True")

For Each objService in colServices
    Wscript.Echo objService.DisplayName 
Next

