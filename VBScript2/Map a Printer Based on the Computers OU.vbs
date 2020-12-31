Set objSysInfo = CreateObject("ADSystemInfo")
strName = objSysInfo.ComputerName

arrComputerName = Split(strName, ",")
arrOU = Split(arrComputerName(1), "=")
strComputerOU = arrOU(1)

Set objNetwork = CreateObject("WScript.Network")

Select Case strComputerOU
    Case "Client"
        objNetwork.AddWindowsPrinterConnection "\\PrintServer1\ClientPrinter"
        objNetwork.SetDefaultPrinter "\\PrintServer1\ClientPrinter"
    Case "Finance"
        objNetwork.AddWindowsPrinterConnection "\\PrintServer2\FinancePrinter"
        objNetwork.SetDefaultPrinter "\\PrintServer2\FinancePrinter"
    Case "Human Resources"
        objNetwork.AddWindowsPrinterConnection "\\PrintServer3\HRPrinter"
        objNetwork.SetDefaultPrinter "\\PrintServer3\HRPrinter"
    Case "Research"
        objNetwork.AddWindowsPrinterConnection "\\PrintServer4\ResearchPrinter"
        objNetwork.SetDefaultPrinter "\\PrintServer4\ResearchPrinter"
    Case Else
        objNetwork.AddWindowsPrinterConnection "\\PrintServer5\GenericPrinter"
        objNetwork.SetDefaultPrinter "\\PrintServer5\GenericPrinter"
End Select
  


