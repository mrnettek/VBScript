' Description: Returns a list of all the services installed on a computer that are currently stopped.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" & _
    "{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2")

Set colStoppedServices = objWMIService.ExecQuery _
    ("Select * From Win32_Service Where State <> 'Running'")
 
For Each objService in colStoppedServices
    Wscript.Echo objService.DisplayName  & " = " & objService.State
Next

