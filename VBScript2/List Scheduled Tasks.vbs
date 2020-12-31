' Description: Enumerates all the scheduled tasks on a computer. Note: WMI can only enumerate scheduled tasks created using the Win32_ScheduledJob class or the At.exe utility. It cannot enumerate tasks created using the Task Scheduler.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colScheduledJobs = objWMIService.ExecQuery _
    ("Select * from Win32_ScheduledJob")

For Each objJob in colScheduledJobs
    Wscript.Echo "Caption: " & objJob.Caption
    Wscript.Echo "Command: " & objJob.Command
    Wscript.Echo "Days of Month: " & objJob.DaysOfMonth
    Wscript.Echo "Days of Week: " & objJob.DaysOfWeek
    Wscript.Echo "Description: " & objJob.Description
    Wscript.Echo "Elapsed Time: " & objJob.ElapsedTime
    Wscript.Echo "Install Date: " & objJob.InstallDate
    Wscript.Echo "Interact with Desktop: " & objJob.InteractWithDesktop
    Wscript.Echo "Job ID: " & objJob.JobID
    Wscript.Echo "Job Status: " & objJob.JobStatus
    Wscript.Echo "Name: " & objJob.Name
    Wscript.Echo "Notify: " & objJob.Notify
    Wscript.Echo "Owner: " & objJob.Owner
    Wscript.Echo "Priority: " & objJob.Priority
    Wscript.Echo "Run Repeatedly: " & objJob.RunRepeatedly
    Wscript.Echo "Start Time: " & objJob.StartTime
    Wscript.Echo "Status: " & objJob.Status
    Wscript.Echo "Time Submitted: " & objJob.TimeSubmitted
    Wscript.Echo "Until Time: " & objJob.UntilTime
Next

