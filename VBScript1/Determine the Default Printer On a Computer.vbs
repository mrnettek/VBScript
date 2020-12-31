strComputer = "."

Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where Default = TRUE")

For Each objPrinter in colPrinters
    Wscript.Echo objPrinter.ShareName
Next
  


