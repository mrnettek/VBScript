' Description: Pauses any printers that have no pending print jobs.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_Printer")

For Each objPrinter in colInstalledPrinters
    Set colPrintJobs = objWMIService.ExecQuery _
        ("Select * from Win32_PerfFormattedData_Spooler_PrintQueue " _
            & "Where Name = '" & objPrinter.Name & "'")
    For Each objPrintQueue in colPrintJobs
        If objPrintQueue.Jobs = 0 and objPrintQueue.Name <> "_Total" Then
            objPrinter.Pause()
        End If
    Next
Next

