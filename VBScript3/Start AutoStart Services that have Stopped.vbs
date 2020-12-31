' Description: Restarts any auto-start services that have been stopped.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colListOfServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where State = 'Stopped' and StartMode = " _
        & "'Auto'")

For Each objService in colListOfServices
    objService.StartService()
Next

