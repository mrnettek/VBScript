strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where Local = TRUE")

Set objNetwork = CreateObject("WScript.Network")
objNetwork.AddWindowsPrinterConnection "\\PrintServer1\Xerox300"

If colPrinters.Count = 0 Then
    objNetwork.SetDefaultPrinter "\\PrintServer1\Xerox300"
End If
  


