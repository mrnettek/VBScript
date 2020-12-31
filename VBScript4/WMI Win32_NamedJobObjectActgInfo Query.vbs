On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NamedJobObjectActgInfo",,48)
For Each objItem in colItems
    Wscript.Echo "ActiveProcesses: " & objItem.ActiveProcesses
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OtherOperationCount: " & objItem.OtherOperationCount
    Wscript.Echo "OtherTransferCount: " & objItem.OtherTransferCount
    Wscript.Echo "PeakJobMemoryUsed: " & objItem.PeakJobMemoryUsed
    Wscript.Echo "PeakProcessMemoryUsed: " & objItem.PeakProcessMemoryUsed
    Wscript.Echo "ReadOperationCount: " & objItem.ReadOperationCount
    Wscript.Echo "ReadTransferCount: " & objItem.ReadTransferCount
    Wscript.Echo "ThisPeriodTotalKernelTime: " & objItem.ThisPeriodTotalKernelTime
    Wscript.Echo "ThisPeriodTotalUserTime: " & objItem.ThisPeriodTotalUserTime
    Wscript.Echo "TotalKernelTime: " & objItem.TotalKernelTime
    Wscript.Echo "TotalPageFaultCount: " & objItem.TotalPageFaultCount
    Wscript.Echo "TotalProcesses: " & objItem.TotalProcesses
    Wscript.Echo "TotalTerminatedProcesses: " & objItem.TotalTerminatedProcesses
    Wscript.Echo "TotalUserTime: " & objItem.TotalUserTime
    Wscript.Echo "WriteOperationCount: " & objItem.WriteOperationCount
    Wscript.Echo "WriteTransferCount: " & objItem.WriteTransferCount
Next

