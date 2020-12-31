' Description: Sets the priority for a printer to 2.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer where DeviceID = 'ArtDepartmentPrinter' ")

For Each objPrinter in colPrinters
    objPrinter.Priority = 2
    objPrinter.Put_
Next

