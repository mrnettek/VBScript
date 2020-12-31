On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_ModifySettingAction",,48)
For Each objItem in colItems
    Wscript.Echo "ActionID: " & objItem.ActionID
    Wscript.Echo "ActionType: " & objItem.ActionType
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Direction: " & objItem.Direction
    Wscript.Echo "EntryName: " & objItem.EntryName
    Wscript.Echo "EntryValue: " & objItem.EntryValue
    Wscript.Echo "FileName: " & objItem.FileName
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SectionKey: " & objItem.SectionKey
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
Next

