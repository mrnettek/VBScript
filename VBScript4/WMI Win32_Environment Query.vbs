On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "SystemVariable: " & objItem.SystemVariable
    Wscript.Echo "UserName: " & objItem.UserName
    Wscript.Echo "VariableValue: " & objItem.VariableValue
Next

