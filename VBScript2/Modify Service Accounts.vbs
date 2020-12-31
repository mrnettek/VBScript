' Description: Changes the service account to LocalService for any services running under the hypothetical service account Netsvc.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServiceList = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where StartName = '.\netsvc'")

For each objService in colServices
    errServiceChange = objService.Change _
        ( , , , , , , "NT AUTHORITY\LocalService" , "")  
Next

