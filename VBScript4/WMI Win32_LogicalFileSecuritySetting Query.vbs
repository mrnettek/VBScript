On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalFileSecuritySetting",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ControlFlags: " & objItem.ControlFlags
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "OwnerPermissions: " & objItem.OwnerPermissions
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo "SettingID: " & objItem.SettingID
Next

