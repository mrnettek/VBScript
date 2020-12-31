On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_Job",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ElapsedTime: " & objItem.ElapsedTime
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "JobStatus: " & objItem.JobStatus
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Notify: " & objItem.Notify
    Wscript.Echo "Owner: " & objItem.Owner
    Wscript.Echo "Priority: " & objItem.Priority
    Wscript.Echo "StartTime: " & objItem.StartTime
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "TimeSubmitted: " & objItem.TimeSubmitted
    Wscript.Echo "UntilTime: " & objItem.UntilTime
Next

