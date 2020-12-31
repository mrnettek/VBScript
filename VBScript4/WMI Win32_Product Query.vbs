On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Product",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "IdentifyingNumber: " & objItem.IdentifyingNumber
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "InstallDate2: " & objItem.InstallDate2
    Wscript.Echo "InstallLocation: " & objItem.InstallLocation
    Wscript.Echo "InstallState: " & objItem.InstallState
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PackageCache: " & objItem.PackageCache
    Wscript.Echo "SKUNumber: " & objItem.SKUNumber
    Wscript.Echo "Vendor: " & objItem.Vendor
    Wscript.Echo "Version: " & objItem.Version
Next

