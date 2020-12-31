strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where DeviceID = 'ArtDepartmentPrinter'")

For Each objPrinter in colPrinters
    objPrinter.KeepPrintedJobs = False
    objPrinter.Put_
Next
  


