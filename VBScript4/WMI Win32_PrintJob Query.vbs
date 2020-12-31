On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PrintJob",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "DataType: " & objItem.DataType
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Document: " & objItem.Document
    Wscript.Echo "DriverName: " & objItem.DriverName
    Wscript.Echo "ElapsedTime: " & objItem.ElapsedTime
    Wscript.Echo "HostPrintQueue: " & objItem.HostPrintQueue
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "JobId: " & objItem.JobId
    Wscript.Echo "JobStatus: " & objItem.JobStatus
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Notify: " & objItem.Notify
    Wscript.Echo "Owner: " & objItem.Owner
    Wscript.Echo "PagesPrinted: " & objItem.PagesPrinted
    Wscript.Echo "Parameters: " & objItem.Parameters
    Wscript.Echo "PrintProcessor: " & objItem.PrintProcessor
    Wscript.Echo "Priority: " & objItem.Priority
    Wscript.Echo "Size: " & objItem.Size
    Wscript.Echo "StartTime: " & objItem.StartTime
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "StatusMask: " & objItem.StatusMask
    Wscript.Echo "TimeSubmitted: " & objItem.TimeSubmitted
    Wscript.Echo "TotalPages: " & objItem.TotalPages
    Wscript.Echo "UntilTime: " & objItem.UntilTime
Next

