On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_SettingCheck",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CheckID: " & objItem.CheckID
    Wscript.Echo "CheckMode: " & objItem.CheckMode
    Wscript.Echo "CheckType: " & objItem.CheckType
    Wscript.Echo "Description: " & objItem.Description
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

