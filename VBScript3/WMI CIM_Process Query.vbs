On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_Process",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CreationDate: " & objItem.CreationDate
    Wscript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ExecutionState: " & objItem.ExecutionState
    Wscript.Echo "Handle: " & objItem.Handle
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "KernelModeTime: " & objItem.KernelModeTime
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OSCreationClassName: " & objItem.OSCreationClassName
    Wscript.Echo "OSName: " & objItem.OSName
    Wscript.Echo "Priority: " & objItem.Priority
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "TerminationDate: " & objItem.TerminationDate
    Wscript.Echo "UserModeTime: " & objItem.UserModeTime
    Wscript.Echo "WorkingSetSize: " & objItem.WorkingSetSize
Next

