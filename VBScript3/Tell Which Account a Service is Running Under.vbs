strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colServices = objWMIService.ExecQuery _
    ("Select * From Win32_Service")
For Each objService in colServices
    Wscript.Echo objService.Name, objService.StartName
Next
  


