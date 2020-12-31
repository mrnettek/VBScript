strComputer = "TestPC"

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery("Select * From Win32_Printer")
     
For Each objPrinter in colPrinters
    If objPrinter.Attributes And 64 Then 
        strPrinterType = "Local"
    Else
        strPrinterType = "Network"
    End If
    Wscript.Echo objPrinter.Name & " -- " & strPrinterType
Next
  


