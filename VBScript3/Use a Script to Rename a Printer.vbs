strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where DeviceID = 'HP'")

For Each objPrinter in colPrinters
    objPrinter.RenamePrinter("ArtDepartmentPrinter")
Next
  


