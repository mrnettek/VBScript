On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_RemoveIniAction",,48)
For Each objItem in colItems
    Wscript.Echo "Action: " & objItem.Action
    Wscript.Echo "ActionID: " & objItem.ActionID
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Direction: " & objItem.Direction
    Wscript.Echo "key: " & objItem.key
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Section: " & objItem.Section
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Value: " & objItem.Value
    Wscript.Echo "Version: " & objItem.Version
Next

