On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PrivilegesStatus",,48)
For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Operation: " & objItem.Operation
    Wscript.Echo "ParameterInfo: " & objItem.ParameterInfo
    Wscript.Echo "PrivilegesNotHeld: " & objItem.PrivilegesNotHeld
    Wscript.Echo "PrivilegesRequired: " & objItem.PrivilegesRequired
    Wscript.Echo "ProviderName: " & objItem.ProviderName
    Wscript.Echo "StatusCode: " & objItem.StatusCode
Next

