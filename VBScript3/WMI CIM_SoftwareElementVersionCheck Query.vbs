On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_SoftwareElementVersionCheck",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CheckID: " & objItem.CheckID
    Wscript.Echo "CheckMode: " & objItem.CheckMode
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "LowerSoftwareElementVersion: " & objItem.LowerSoftwareElementVersion
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementName: " & objItem.SoftwareElementName
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "SoftwareElementStateDesired: " & objItem.SoftwareElementStateDesired
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "TargetOperatingSystemDesired: " & objItem.TargetOperatingSystemDesired
    Wscript.Echo "UpperSoftwareElementVersion: " & objItem.UpperSoftwareElementVersion
    Wscript.Echo "Version: " & objItem.Version
Next

