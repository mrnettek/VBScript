On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ServiceSpecification",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CheckID: " & objItem.CheckID
    Wscript.Echo "CheckMode: " & objItem.CheckMode
    Wscript.Echo "Dependencies: " & objItem.Dependencies
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DisplayName: " & objItem.DisplayName
    Wscript.Echo "ErrorControl: " & objItem.ErrorControl
    Wscript.Echo "ID: " & objItem.ID
    Wscript.Echo "LoadOrderGroup: " & objItem.LoadOrderGroup
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Password: " & objItem.Password
    Wscript.Echo "ServiceType: " & objItem.ServiceType
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "StartName: " & objItem.StartName
    Wscript.Echo "StartType: " & objItem.StartType
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
Next

