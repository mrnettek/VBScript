On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSAccount",,48)
For Each objItem in colItems
    Wscript.Echo "AccountName: " & objItem.AccountName
    Wscript.Echo "AuditFail: " & objItem.AuditFail
    Wscript.Echo "AuditSuccess: " & objItem.AuditSuccess
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "PermissionsAllowed: " & objItem.PermissionsAllowed
    Wscript.Echo "PermissionsDenied: " & objItem.PermissionsDenied
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "TerminalName: " & objItem.TerminalName
Next

