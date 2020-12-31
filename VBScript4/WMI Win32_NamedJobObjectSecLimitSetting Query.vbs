On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NamedJobObjectSecLimitSetting",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "PrivilegesToDelete: " & objItem.PrivilegesToDelete
    Wscript.Echo "RestrictedSIDs: " & objItem.RestrictedSIDs
    Wscript.Echo "SecurityLimitFlags: " & objItem.SecurityLimitFlags
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "SIDsToDisable: " & objItem.SIDsToDisable
Next

