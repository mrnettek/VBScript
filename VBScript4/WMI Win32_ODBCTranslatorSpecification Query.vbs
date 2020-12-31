On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ODBCTranslatorSpecification",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CheckID: " & objItem.CheckID
    Wscript.Echo "CheckMode: " & objItem.CheckMode
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "File: " & objItem.File
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SetupFile: " & objItem.SetupFile
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Translator: " & objItem.Translator
    Wscript.Echo "Version: " & objItem.Version
Next

