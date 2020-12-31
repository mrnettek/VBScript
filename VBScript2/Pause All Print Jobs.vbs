' Description: Pauses all the print jobs on a print server.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrintJobs =  objWMIService.ExecQuery _
    ("Select * from Win32_PrintJob")

For Each objPrintJob in colPrintJobs 
    objPrintJob.Pause
Next

