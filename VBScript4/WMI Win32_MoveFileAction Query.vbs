On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_MoveFileAction",,48)
For Each objItem in colItems
    Wscript.Echo "ActionID: " & objItem.ActionID
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DestFolder: " & objItem.DestFolder
    Wscript.Echo "DestName: " & objItem.DestName
    Wscript.Echo "Direction: " & objItem.Direction
    Wscript.Echo "FileKey: " & objItem.FileKey
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Options: " & objItem.Options
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "SourceFolder: " & objItem.SourceFolder
    Wscript.Echo "SourceName: " & objItem.SourceName
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
Next

