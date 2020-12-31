On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_JobObjectStatus",,48)
For Each objItem in colItems
    Wscript.Echo "AdditionalDescription: " & objItem.AdditionalDescription
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Operation: " & objItem.Operation
    Wscript.Echo "ParameterInfo: " & objItem.ParameterInfo
    Wscript.Echo "ProviderName: " & objItem.ProviderName
    Wscript.Echo "StatusCode: " & objItem.StatusCode
    Wscript.Echo "Win32ErrorCode: " & objItem.Win32ErrorCode
Next

