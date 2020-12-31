On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ShortcutAction",,48)
For Each objItem in colItems
    Wscript.Echo "ActionID: " & objItem.ActionID
    Wscript.Echo "Arguments: " & objItem.Arguments
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Direction: " & objItem.Direction
    Wscript.Echo "HotKey: " & objItem.HotKey
    Wscript.Echo "IconIndex: " & objItem.IconIndex
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Shortcut: " & objItem.Shortcut
    Wscript.Echo "ShowCmd: " & objItem.ShowCmd
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "Target: " & objItem.Target
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "WkDir: " & objItem.WkDir
Next

