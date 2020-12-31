On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ProgramGroup",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "GroupName: " & objItem.GroupName
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "UserName: " & objItem.UserName
Next

