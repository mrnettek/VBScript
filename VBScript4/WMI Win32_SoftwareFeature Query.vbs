On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_SoftwareFeature",,48)
For Each objItem in colItems
    Wscript.Echo "Accesses: " & objItem.Accesses
    Wscript.Echo "Attributes: " & objItem.Attributes
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "IdentifyingNumber: " & objItem.IdentifyingNumber
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "InstallState: " & objItem.InstallState
    Wscript.Echo "LastUse: " & objItem.LastUse
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ProductName: " & objItem.ProductName
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "Vendor: " & objItem.Vendor
    Wscript.Echo "Version: " & objItem.Version
Next

