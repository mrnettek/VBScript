On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TokenPrivileges",,48)
For Each objItem in colItems
    Wscript.Echo "PrivilegeCount: " & objItem.PrivilegeCount
    Wscript.Echo "Privileges: " & objItem.Privileges
Next

