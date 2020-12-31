On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ServerSession",,48)
For Each objItem in colItems
    Wscript.Echo "ActiveTime: " & objItem.ActiveTime
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ClientType: " & objItem.ClientType
    Wscript.Echo "ComputerName: " & objItem.ComputerName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "IdleTime: " & objItem.IdleTime
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ResourcesOpened: " & objItem.ResourcesOpened
    Wscript.Echo "SessionType: " & objItem.SessionType
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "TransportName: " & objItem.TransportName
    Wscript.Echo "UserName: " & objItem.UserName
Next

