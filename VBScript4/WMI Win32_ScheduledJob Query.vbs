On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ScheduledJob",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Command: " & objItem.Command
    Wscript.Echo "DaysOfMonth: " & objItem.DaysOfMonth
    Wscript.Echo "DaysOfWeek: " & objItem.DaysOfWeek
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ElapsedTime: " & objItem.ElapsedTime
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "InteractWithDesktop: " & objItem.InteractWithDesktop
    Wscript.Echo "JobId: " & objItem.JobId
    Wscript.Echo "JobStatus: " & objItem.JobStatus
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Notify: " & objItem.Notify
    Wscript.Echo "Owner: " & objItem.Owner
    Wscript.Echo "Priority: " & objItem.Priority
    Wscript.Echo "RunRepeatedly: " & objItem.RunRepeatedly
    Wscript.Echo "StartTime: " & objItem.StartTime
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "TimeSubmitted: " & objItem.TimeSubmitted
    Wscript.Echo "UntilTime: " & objItem.UntilTime
Next

