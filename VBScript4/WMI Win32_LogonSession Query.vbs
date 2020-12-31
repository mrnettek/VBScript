On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogonSession",,48)
For Each objItem in colItems
    Wscript.Echo "AuthenticationPackage: " & objItem.AuthenticationPackage
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LogonId: " & objItem.LogonId
    Wscript.Echo "LogonType: " & objItem.LogonType
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "StartTime: " & objItem.StartTime
    Wscript.Echo "Status: " & objItem.Status
Next

