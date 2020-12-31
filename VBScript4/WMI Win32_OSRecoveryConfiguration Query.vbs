On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OSRecoveryConfiguration",,48)
For Each objItem in colItems
    Wscript.Echo "AutoReboot: " & objItem.AutoReboot
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "DebugFilePath: " & objItem.DebugFilePath
    Wscript.Echo "DebugInfoType: " & objItem.DebugInfoType
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ExpandedDebugFilePath: " & objItem.ExpandedDebugFilePath
    Wscript.Echo "ExpandedMiniDumpDirectory: " & objItem.ExpandedMiniDumpDirectory
    Wscript.Echo "KernelDumpOnly: " & objItem.KernelDumpOnly
    Wscript.Echo "MiniDumpDirectory: " & objItem.MiniDumpDirectory
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OverwriteExistingDebugFile: " & objItem.OverwriteExistingDebugFile
    Wscript.Echo "SendAdminAlert: " & objItem.SendAdminAlert
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "WriteDebugInfo: " & objItem.WriteDebugInfo
    Wscript.Echo "WriteToSystemLog: " & objItem.WriteToSystemLog
Next

