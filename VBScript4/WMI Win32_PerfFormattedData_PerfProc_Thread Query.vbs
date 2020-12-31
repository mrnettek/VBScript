On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfProc_Thread",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ContextSwitchesPersec: " & objItem.ContextSwitchesPersec
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ElapsedTime: " & objItem.ElapsedTime
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "IDProcess: " & objItem.IDProcess
    Wscript.Echo "IDThread: " & objItem.IDThread
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PercentPrivilegedTime: " & objItem.PercentPrivilegedTime
    Wscript.Echo "PercentProcessorTime: " & objItem.PercentProcessorTime
    Wscript.Echo "PercentUserTime: " & objItem.PercentUserTime
    Wscript.Echo "PriorityBase: " & objItem.PriorityBase
    Wscript.Echo "PriorityCurrent: " & objItem.PriorityCurrent
    Wscript.Echo "StartAddress: " & objItem.StartAddress
    Wscript.Echo "ThreadState: " & objItem.ThreadState
    Wscript.Echo "ThreadWaitReason: " & objItem.ThreadWaitReason
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

