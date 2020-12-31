On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TypeLibraryAction",,48)
For Each objItem in colItems
    Wscript.Echo "ActionID: " & objItem.ActionID
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Cost: " & objItem.Cost
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Direction: " & objItem.Direction
    Wscript.Echo "Language: " & objItem.Language
    Wscript.Echo "LibID: " & objItem.LibID
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
Next

