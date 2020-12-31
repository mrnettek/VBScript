On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DuplicateFileAction",,48)
For Each objItem in colItems
    Wscript.Echo "ActionID: " & objItem.ActionID
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "DeleteAfterCopy: " & objItem.DeleteAfterCopy
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Destination: " & objItem.Destination
    Wscript.Echo "Direction: " & objItem.Direction
    Wscript.Echo "FileKey: " & objItem.FileKey
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "Source: " & objItem.Source
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
Next

