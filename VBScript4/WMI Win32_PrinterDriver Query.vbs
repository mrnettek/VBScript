On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PrinterDriver",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ConfigFile: " & objItem.ConfigFile
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "DataFile: " & objItem.DataFile
    Wscript.Echo "DefaultDataType: " & objItem.DefaultDataType
    Wscript.Echo "DependentFiles: " & objItem.DependentFiles
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DriverPath: " & objItem.DriverPath
    Wscript.Echo "FilePath: " & objItem.FilePath
    Wscript.Echo "HelpFile: " & objItem.HelpFile
    Wscript.Echo "InfName: " & objItem.InfName
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "MonitorName: " & objItem.MonitorName
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OEMUrl: " & objItem.OEMUrl
    Wscript.Echo "Started: " & objItem.Started
    Wscript.Echo "StartMode: " & objItem.StartMode
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "SupportedPlatform: " & objItem.SupportedPlatform
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "Version: " & objItem.Version
Next

