On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_FileSpecification",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CheckID: " & objItem.CheckID
    Wscript.Echo "CheckMode: " & objItem.CheckMode
    Wscript.Echo "CheckSum: " & objItem.CheckSum
    Wscript.Echo "CRC1: " & objItem.CRC1
    Wscript.Echo "CRC2: " & objItem.CRC2
    Wscript.Echo "CreateTimeStamp: " & objItem.CreateTimeStamp
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "FileSize: " & objItem.FileSize
    Wscript.Echo "MD5Checksum: " & objItem.MD5Checksum
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SoftwareElementID: " & objItem.SoftwareElementID
    Wscript.Echo "SoftwareElementState: " & objItem.SoftwareElementState
    Wscript.Echo "TargetOperatingSystem: " & objItem.TargetOperatingSystem
    Wscript.Echo "Version: " & objItem.Version
Next

