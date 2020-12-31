' Description: Schedules Autochk.exe to run against drive C the next time the computer reboots.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objDisk = objWMIService.Get("Win32_LogicalDisk")

errReturn = objDisk.ScheduleAutoChk(Array("C:"))

