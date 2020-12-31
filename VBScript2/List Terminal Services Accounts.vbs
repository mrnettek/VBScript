' Description: Returns information about all the Terminal Services accounts on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSAccount")

For Each objItem in colItems
    Wscript.Echo "Account name: " & objItem.AccountName
    Wscript.Echo "Audit failure: " & objItem.AuditFail
    Wscript.Echo "Audit success: " & objItem.AuditSuccess
    Wscript.Echo "Permissions allowed: " & objItem.PermissionsAllowed
    Wscript.Echo "Permissions denied: " & objItem.PermissionsDenied
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "Terminal name: " & objItem.TerminalName
    Wscript.Echo
Next

