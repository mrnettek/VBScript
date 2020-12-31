On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_OperatingSystem",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "CurrentTimeZone: " & objItem.CurrentTimeZone
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Distributed: " & objItem.Distributed
    Wscript.Echo "FreePhysicalMemory: " & objItem.FreePhysicalMemory
    Wscript.Echo "FreeSpaceInPagingFiles: " & objItem.FreeSpaceInPagingFiles
    Wscript.Echo "FreeVirtualMemory: " & objItem.FreeVirtualMemory
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LastBootUpTime: " & objItem.LastBootUpTime
    Wscript.Echo "LocalDateTime: " & objItem.LocalDateTime
    Wscript.Echo "MaxNumberOfProcesses: " & objItem.MaxNumberOfProcesses
    Wscript.Echo "MaxProcessMemorySize: " & objItem.MaxProcessMemorySize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfLicensedUsers: " & objItem.NumberOfLicensedUsers
    Wscript.Echo "NumberOfProcesses: " & objItem.NumberOfProcesses
    Wscript.Echo "NumberOfUsers: " & objItem.NumberOfUsers
    Wscript.Echo "OSType: " & objItem.OSType
    Wscript.Echo "OtherTypeDescription: " & objItem.OtherTypeDescription
    Wscript.Echo "SizeStoredInPagingFiles: " & objItem.SizeStoredInPagingFiles
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "TotalSwapSpaceSize: " & objItem.TotalSwapSpaceSize
    Wscript.Echo "TotalVirtualMemorySize: " & objItem.TotalVirtualMemorySize
    Wscript.Echo "TotalVisibleMemorySize: " & objItem.TotalVisibleMemorySize
    Wscript.Echo "Version: " & objItem.Version
Next

