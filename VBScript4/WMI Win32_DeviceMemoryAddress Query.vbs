On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DeviceMemoryAddress",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "EndingAddress: " & objItem.EndingAddress
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "MemoryType: " & objItem.MemoryType
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "StartingAddress: " & objItem.StartingAddress
    Wscript.Echo "Status: " & objItem.Status
Next

