On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "IdentifyingNumber: " & objItem.IdentifyingNumber
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SKUNumber: " & objItem.SKUNumber
    Wscript.Echo "UUID: " & objItem.UUID
    Wscript.Echo "Vendor: " & objItem.Vendor
    Wscript.Echo "Version: " & objItem.Version
Next

