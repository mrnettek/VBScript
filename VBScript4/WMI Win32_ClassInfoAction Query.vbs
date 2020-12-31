On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ClassInfoAction",,48)
For Each objItem in colItems
    Wscript.Echo "ActionID: " & objItem.ActionID
    Wscript.Echo "AppID: " & objItem.AppID
    Wscript.Echo "Argument: " & objItem.Argument
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CLSID: " & objItem.CLSID
    Wscript.Echo "Context: " & objItem.Context
    Wscript.Echo "DefInprocHandler: " & objItem.DefInprocHandler
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Direction: " & objItem.Direction
    Wscript.Echo "FileTypeMask: " & objItem.FileTypeMask
    Wscript.Echo "Insertable: " & objItem.Insertable
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ProgID: " & objItem.ProgID
    Wscript.Echo "RemoteName: " & objItem.RemoteName
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "VIProgID: " & objItem.VIProgID
Next

