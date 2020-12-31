' Description: Reports processor use time, in seconds, for each process running on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")

For Each objProcess in colProcesses
    sngProcessTime = (CSng(objProcess.KernelModeTime) + _
            CSng(objProcess.UserModeTime)) / 10000000
    Wscript.Echo objProcess.name & VbTab & sngProcessTime
Next

