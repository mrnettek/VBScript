strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery _
    ("Select * from Win32_PrinterConfiguration Where Color = 1")

For Each objPrinter in colPrinters
    strName = objPrinter.Name
    strName = Replace(strName, "\", "\\")

    Set colBWPrinters = objWMIService.ExecQuery _
        ("Select * from Win32_Printer Where Name = '" & strName & "'")
    For Each objBWPrinter in colBWPrinters
        objBWSPrinter.Delete_
    Next
Next
  


