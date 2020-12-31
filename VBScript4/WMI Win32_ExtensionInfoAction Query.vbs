On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ExtensionInfoAction",,48)
For Each objItem in colItems
    Wscript.Echo "ActionID: " & objItem.ActionID
    Wscript.Echo "Argument: " & objItem.Argument
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Command: " & objItem.Command
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Direction: " & objItem.Direction
    Wscript.Echo "Extension: " & objItem.Extension
    Wscript.Echo "MIME: " & objItem.MIME
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ProgID: " & objItem.ProgID
    Wscript.Echo "ShellNew: " & objItem.ShellNew
    Wscript.Echo "ShellNewValue: " & objItem.ShellNewValue
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Verb: " & objItem.Verb
    Wscript.Echo "Version: " & objItem.Version
Next

