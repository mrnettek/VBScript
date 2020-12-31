On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ServerConnection",,48)
For Each objItem in colItems
    Wscript.Echo "ActiveTime: " & objItem.ActiveTime
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ComputerName: " & objItem.ComputerName
    Wscript.Echo "ConnectionID: " & objItem.ConnectionID
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfFiles: " & objItem.NumberOfFiles
    Wscript.Echo "NumberOfUsers: " & objItem.NumberOfUsers
    Wscript.Echo "ShareName: " & objItem.ShareName
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "UserName: " & objItem.UserName
Next

