On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPSignedDriver",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ClassGuid: " & objItem.ClassGuid
    Wscript.Echo "CompatID: " & objItem.CompatID
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DeviceClass: " & objItem.DeviceClass
    Wscript.Echo "DeviceID: " & objItem.DeviceID
    Wscript.Echo "DeviceName: " & objItem.DeviceName
    Wscript.Echo "DevLoader: " & objItem.DevLoader
    Wscript.Echo "DriverDate: " & objItem.DriverDate
    Wscript.Echo "DriverName: " & objItem.DriverName
    Wscript.Echo "DriverProviderName: " & objItem.DriverProviderName
    Wscript.Echo "DriverVersion: " & objItem.DriverVersion
    Wscript.Echo "FriendlyName: " & objItem.FriendlyName
    Wscript.Echo "HardWareID: " & objItem.HardWareID
    Wscript.Echo "InfName: " & objItem.InfName
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "IsSigned: " & objItem.IsSigned
    Wscript.Echo "Location: " & objItem.Location
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PDO: " & objItem.PDO
    Wscript.Echo "Signer: " & objItem.Signer
    Wscript.Echo "Started: " & objItem.Started
    Wscript.Echo "StartMode: " & objItem.StartMode
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "SystemCreationClassName: " & objItem.SystemCreationClassName
    Wscript.Echo "SystemName: " & objItem.SystemName
Next

