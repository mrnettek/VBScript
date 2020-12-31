' Description: Reports statistics such as thread count and working set size for all the processes running on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")

For Each objProcess in colProcessList
    Wscript.Echo "Process: " & objProcess.Name 
    Wscript.Echo "Process ID: " & objProcess.ProcessID 
    Wscript.Echo "Thread Count: " & objProcess.ThreadCount 
    Wscript.Echo "Page File Size: " & objProcess.PageFileUsage 
    Wscript.Echo "Page Faults: " & objProcess.PageFaults 
    Wscript.Echo "Working Set Size: " & objProcess.WorkingSetSize 
Next

