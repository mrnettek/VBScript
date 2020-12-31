On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
For Each objItem in colItems
    Wscript.Echo "BootDevice: " & objItem.BootDevice
    Wscript.Echo "BuildNumber: " & objItem.BuildNumber
    Wscript.Echo "BuildType: " & objItem.BuildType
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CodeSet: " & objItem.CodeSet
    Wscript.Echo "CountryCode: " & objItem.CountryCode
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
    Wscript.Echo "CSDVersion: " & objItem.CSDVersion
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "CurrentTimeZone: " & objItem.CurrentTimeZone
    Wscript.Echo "DataExecutionPrevention_32BitApplications: " & objItem.DataExecutionPrevention_32BitApplications
    Wscript.Echo "DataExecutionPrevention_Available: " & objItem.DataExecutionPrevention_Available
    Wscript.Echo "DataExecutionPrevention_Drivers: " & objItem.DataExecutionPrevention_Drivers
    Wscript.Echo "DataExecutionPrevention_SupportPolicy: " & objItem.DataExecutionPrevention_SupportPolicy
    Wscript.Echo "Debug: " & objItem.Debug
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Distributed: " & objItem.Distributed
    Wscript.Echo "EncryptionLevel: " & objItem.EncryptionLevel
    Wscript.Echo "ForegroundApplicationBoost: " & objItem.ForegroundApplicationBoost
    Wscript.Echo "FreePhysicalMemory: " & objItem.FreePhysicalMemory
    Wscript.Echo "FreeSpaceInPagingFiles: " & objItem.FreeSpaceInPagingFiles
    Wscript.Echo "FreeVirtualMemory: " & objItem.FreeVirtualMemory
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LargeSystemCache: " & objItem.LargeSystemCache
    Wscript.Echo "LastBootUpTime: " & objItem.LastBootUpTime
    Wscript.Echo "LocalDateTime: " & objItem.LocalDateTime
    Wscript.Echo "Locale: " & objItem.Locale
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "MaxNumberOfProcesses: " & objItem.MaxNumberOfProcesses
    Wscript.Echo "MaxProcessMemorySize: " & objItem.MaxProcessMemorySize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfLicensedUsers: " & objItem.NumberOfLicensedUsers
    Wscript.Echo "NumberOfProcesses: " & objItem.NumberOfProcesses
    Wscript.Echo "NumberOfUsers: " & objItem.NumberOfUsers
    Wscript.Echo "Organization: " & objItem.Organization
    Wscript.Echo "OSLanguage: " & objItem.OSLanguage
    Wscript.Echo "OSProductSuite: " & objItem.OSProductSuite
    Wscript.Echo "OSType: " & objItem.OSType
    Wscript.Echo "OtherTypeDescription: " & objItem.OtherTypeDescription
    Wscript.Echo "PlusProductID: " & objItem.PlusProductID
    Wscript.Echo "PlusVersionNumber: " & objItem.PlusVersionNumber
    Wscript.Echo "Primary: " & objItem.Primary
    Wscript.Echo "ProductType: " & objItem.ProductType
    Wscript.Echo "QuantumLength: " & objItem.QuantumLength
    Wscript.Echo "QuantumType: " & objItem.QuantumType
    Wscript.Echo "RegisteredUser: " & objItem.RegisteredUser
    Wscript.Echo "SerialNumber: " & objItem.SerialNumber
    Wscript.Echo "ServicePackMajorVersion: " & objItem.ServicePackMajorVersion
    Wscript.Echo "ServicePackMinorVersion: " & objItem.ServicePackMinorVersion
    Wscript.Echo "SizeStoredInPagingFiles: " & objItem.SizeStoredInPagingFiles
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "SuiteMask: " & objItem.SuiteMask
    Wscript.Echo "SystemDevice: " & objItem.SystemDevice
    Wscript.Echo "SystemDirectory: " & objItem.SystemDirectory
    Wscript.Echo "SystemDrive: " & objItem.SystemDrive
    Wscript.Echo "TotalSwapSpaceSize: " & objItem.TotalSwapSpaceSize
    Wscript.Echo "TotalVirtualMemorySize: " & objItem.TotalVirtualMemorySize
    Wscript.Echo "TotalVisibleMemorySize: " & objItem.TotalVisibleMemorySize
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "WindowsDirectory: " & objItem.WindowsDirectory
Next

