On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSSessionDirectoryError",,48)
For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Operation: " & objItem.Operation
    Wscript.Echo "ParameterInfo: " & objItem.ParameterInfo
    Wscript.Echo "ProviderName: " & objItem.ProviderName
    Wscript.Echo "SessionDirectory: " & objItem.SessionDirectory
    Wscript.Echo "StatusCode: " & objItem.StatusCode
Next

