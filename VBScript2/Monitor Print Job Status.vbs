' Description: Returns the job ID, user name, and total pages for each print job on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrintJobs =  objWMIService.ExecQuery _
    ("Select * from Win32_PrintJob")

Wscript.Echo "Print Queue, Job ID, Owner, Total Pages"

For Each objPrintJob in colPrintJobs
    strPrinter = Split(objPrintJob.Name,",",-1,1)
    Wscript.Echo strPrinter(0) & ", " & _
        objPrintJob.JobID & ", " &  objPrintJob.Owner & ", " _
            & objPrintJob.TotalPages
Next

