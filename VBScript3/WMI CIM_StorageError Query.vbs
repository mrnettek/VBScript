On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_StorageError",,48)
For Each objItem in colItems
    Wscript.Echo "DeviceCreationClassName: " & objItem.DeviceCreationClassName
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "EndingAddress: " & objItem.EndingAddress
    Wscript.Echo "StartingAddress: " & objItem.StartingAddress
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
Next

