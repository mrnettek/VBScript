On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_StartupCommand",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Command: " & objItem.Command
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Location: " & objItem.Location
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "User: " & objItem.User
Next

