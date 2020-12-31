' Description: Checks the status for each printer on a computer, and issues an alert if any of these printers have stopped.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer Where PrinterStatus = '1' " _
        & "or PrinterStatus = '2'")

If colInstalledPrinters.Count = 0 Then
    Wscript.Echo "All printers are functioning correctly."
Else
    For Each objPrinter in colInstalledPrinters
        Wscript.Echo "Printer " & objprinter.Name & " is not responding." 
    Next
End If

