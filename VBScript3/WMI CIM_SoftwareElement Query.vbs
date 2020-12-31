On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_SoftwareElement",,48)
For Each objItem in colItems
    Wscript.Echo "BuildNumber: " & objItem.BuildNumber
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CodeSet: " & objItem.CodeSet
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "IdentificationCode: " & objItem.IdentificationCode
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "LanguageEdition: " & objItem.LanguageEdition
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OtherTargetOS: " & objItem.OtherTargetOS
    Wscript.Echo "SerialNumber: " & objItem.SerialNumber
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
Next

