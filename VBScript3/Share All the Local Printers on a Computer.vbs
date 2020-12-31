strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where Network = FALSE")

i = 1

For Each objPrinter in colInstalledPrinters
    objPrinter.Shared = TRUE
    objPrinter.ShareName = "Printer" & i
    objPrinter.Put_
    i = i + 1
Next
  


