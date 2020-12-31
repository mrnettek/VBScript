' Description: Configures all auto-start services to issue an alert if the service fails during startup.


Const NORMAL_ERROR_CONTROL = 2

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServiceList = objWMIService.ExecQuery _
    ("Select * from Win32_Service where ErrorControl = 'Ignore'")

For Each objService in colServiceList
    errReturn = objService.Change( , , , NORMAL_ERROR_CONTROL)   
Next

