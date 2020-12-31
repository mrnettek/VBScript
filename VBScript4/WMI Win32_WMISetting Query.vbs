On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_WMISetting",,48)
For Each objItem in colItems
    Wscript.Echo "ASPScriptDefaultNamespace: " & objItem.ASPScriptDefaultNamespace
    Wscript.Echo "ASPScriptEnabled: " & objItem.ASPScriptEnabled
    Wscript.Echo "AutorecoverMofs: " & objItem.AutorecoverMofs
    Wscript.Echo "AutoStartWin9X: " & objItem.AutoStartWin9X
    Wscript.Echo "BackupInterval: " & objItem.BackupInterval
    Wscript.Echo "BackupLastTime: " & objItem.BackupLastTime
    Wscript.Echo "BuildVersion: " & objItem.BuildVersion
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "DatabaseDirectory: " & objItem.DatabaseDirectory
    Wscript.Echo "DatabaseMaxSize: " & objItem.DatabaseMaxSize
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "EnableAnonWin9xConnections: " & objItem.EnableAnonWin9xConnections
    Wscript.Echo "EnableEvents: " & objItem.EnableEvents
    Wscript.Echo "EnableStartupHeapPreallocation: " & objItem.EnableStartupHeapPreallocation
    Wscript.Echo "HighThresholdOnClientObjects: " & objItem.HighThresholdOnClientObjects
    Wscript.Echo "HighThresholdOnEvents: " & objItem.HighThresholdOnEvents
    Wscript.Echo "InstallationDirectory: " & objItem.InstallationDirectory
    Wscript.Echo "LastStartupHeapPreallocation: " & objItem.LastStartupHeapPreallocation
    Wscript.Echo "LoggingDirectory: " & objItem.LoggingDirectory
    Wscript.Echo "LoggingLevel: " & objItem.LoggingLevel
    Wscript.Echo "LowThresholdOnClientObjects: " & objItem.LowThresholdOnClientObjects
    Wscript.Echo "LowThresholdOnEvents: " & objItem.LowThresholdOnEvents
    Wscript.Echo "MaxLogFileSize: " & objItem.MaxLogFileSize
    Wscript.Echo "MaxWaitOnClientObjects: " & objItem.MaxWaitOnClientObjects
    Wscript.Echo "MaxWaitOnEvents: " & objItem.MaxWaitOnEvents
    Wscript.Echo "MofSelfInstallDirectory: " & objItem.MofSelfInstallDirectory
    Wscript.Echo "SettingID: " & objItem.SettingID
Next

