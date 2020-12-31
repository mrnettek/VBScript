On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NamedJobObjectLimitSetting",,48)
For Each objItem in colItems
    Wscript.Echo "ActiveProcessLimit: " & objItem.ActiveProcessLimit
    Wscript.Echo "Affinity: " & objItem.Affinity
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "JobMemoryLimit: " & objItem.JobMemoryLimit
    Wscript.Echo "LimitFlags: " & objItem.LimitFlags
    Wscript.Echo "MaximumWorkingSetSize: " & objItem.MaximumWorkingSetSize
    Wscript.Echo "MinimumWorkingSetSize: " & objItem.MinimumWorkingSetSize
    Wscript.Echo "PerJobUserTimeLimit: " & objItem.PerJobUserTimeLimit
    Wscript.Echo "PerProcessUserTimeLimit: " & objItem.PerProcessUserTimeLimit
    Wscript.Echo "PriorityClass: " & objItem.PriorityClass
    Wscript.Echo "ProcessMemoryLimit: " & objItem.ProcessMemoryLimit
    Wscript.Echo "SchedulingClass: " & objItem.SchedulingClass
    Wscript.Echo "SettingID: " & objItem.SettingID
Next

