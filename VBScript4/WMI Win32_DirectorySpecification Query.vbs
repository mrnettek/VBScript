On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DirectorySpecification",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CheckID: " & objItem.CheckID
    Wscript.Echo "CheckMode: " & objItem.CheckMode
    Wscript.Echo "DefaultDir: " & objItem.DefaultDir
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Directory: " & objItem.Directory
    Wscript.Echo "DirectoryPath: " & objItem.DirectoryPath
    Wscript.Echo "DirectoryType: " & objItem.DirectoryType
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
Next

