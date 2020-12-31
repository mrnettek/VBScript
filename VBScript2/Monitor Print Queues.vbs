' Description: Uses cooked performance counters to return the number of jobs currently in each print queue on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrintQueues =  objWMIService.ExecQuery _
    ("Select * from Win32_PerfFormattedData_Spooler_PrintQueue " & _
        "Where Name <> '_Total'")

For Each objPrintQueue in colPrintQueues
    Wscript.Echo "Name: " & objPrintQueue.Name
    Wscript.Echo "Current jobs: " & objPrintQueue.Jobs
Next

