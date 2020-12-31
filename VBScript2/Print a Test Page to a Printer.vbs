strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where DeviceID = '\\\\atl-ps-01\\color-printer'")

For Each objPrinter in colPrinters
    errReturn = objPrinter.PrintTestPage
    If errReturn = 0 Then
        Wscript.Echo "The test page was printed successfully."
    Else
        Wscript.Echo "The test page could not be printed."
    End If
Next
  


