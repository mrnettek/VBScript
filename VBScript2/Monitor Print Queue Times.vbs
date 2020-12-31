' Description: Identifies any print jobs that have been in the print queue for more than 15 minutes.


Const USE_LOCAL_TIME = True

Set DateTime = CreateObject("WbemScripting.SWbemDateTime")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colInstalledPrinters =  objWMIService.ExecQuery _
    ("Select * from Win32_PrintJob")

Wscript.Echo "Print Queue, Job ID, TimeSubmitted, Total Pages"

For Each objPrinter in colInstalledPrinters
    DateTime.Value = objPrinter.TimeSubmitted
    dtmActualTime = DateTime.GetVarDate(USE_LOCAL_TIME)
    TimeinQueue = DateDiff("n", actualTime, Now)
    If TimeinQueue > 15 Then
        strPrinterName = Split(objPrinter.Name,",",-1,1)
        Wscript.Echo strPrinterName(0) & ", " _
            & objPrinter.JobID & ", " & dtmActualTime & ", " & _
                objPrinter.TotalPages
    End If
Next

