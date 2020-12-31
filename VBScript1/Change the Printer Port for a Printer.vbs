strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where DeviceID='Art Department Printer'")

For Each objPrinter in colPrinters
    objPrinter.PortName = "LPT1:"
    objPrinter.Put_
Next
  


