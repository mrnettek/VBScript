' Description: Uses WMI to return information about all the environment variables on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment")

For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "System Variable: " & objItem.SystemVariable
    Wscript.Echo "User Name: " & objItem.UserName
    Wscript.Echo "Variable Value: " & objItem.VariableValue
Next

