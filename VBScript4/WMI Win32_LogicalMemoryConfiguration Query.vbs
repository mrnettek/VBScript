On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalMemoryConfiguration",,48)
For Each objItem in colItems
    Wscript.Echo "AvailableVirtualMemory: " & objItem.AvailableVirtualMemory
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "TotalPageFileSpace: " & objItem.TotalPageFileSpace
    Wscript.Echo "TotalPhysicalMemory: " & objItem.TotalPhysicalMemory
    Wscript.Echo "TotalVirtualMemory: " & objItem.TotalVirtualMemory
Next

