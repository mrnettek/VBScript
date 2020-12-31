' Description: Changes the service account password for any services running under the hypothetical service account Netsvc.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServiceList = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where StartName = '.\netsvc'")

For Each objService in colServiceList
    errReturn = objService.Change( , , , , , , , "password")  
Next

