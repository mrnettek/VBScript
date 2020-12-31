On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Thread",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ElapsedTime: " & objItem.ElapsedTime
    Wscript.Echo "ExecutionState: " & objItem.ExecutionState
    Wscript.Echo "Handle: " & objItem.Handle
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "KernelModeTime: " & objItem.KernelModeTime
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OSCreationClassName: " & objItem.OSCreationClassName
    Wscript.Echo "OSName: " & objItem.OSName
    Wscript.Echo "Priority: " & objItem.Priority
    Wscript.Echo "PriorityBase: " & objItem.PriorityBase
    Wscript.Echo "ProcessCreationClassName: " & objItem.ProcessCreationClassName
    Wscript.Echo "ProcessHandle: " & objItem.ProcessHandle
    Wscript.Echo "StartAddress: " & objItem.StartAddress
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "ThreadState: " & objItem.ThreadState
    Wscript.Echo "ThreadWaitReason: " & objItem.ThreadWaitReason
    Wscript.Echo "UserModeTime: " & objItem.UserModeTime
Next

