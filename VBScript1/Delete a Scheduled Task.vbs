' Description: Deletes the scheduled task with the Job ID of 1.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objInstance = objWMIService.Get("Win32_ScheduledJob.JobID=1")

err = objInstance.Delete

