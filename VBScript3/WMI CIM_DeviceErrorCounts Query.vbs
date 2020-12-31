On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_DeviceErrorCounts",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CriticalErrorCount: " & objItem.CriticalErrorCount
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceCreationClassName: " & objItem.DeviceCreationClassName
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "IndeterminateErrorCount: " & objItem.IndeterminateErrorCount
    Wscript.Echo "MajorErrorCount: " & objItem.MajorErrorCount
    Wscript.Echo "MinorErrorCount: " & objItem.MinorErrorCount
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
    Wscript.Echo "WarningCount: " & objItem.WarningCount
Next

