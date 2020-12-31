On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_FileSystem",,48)
For Each objItem in colItems
    Wscript.Echo "AvailableSpace: " & objItem.AvailableSpace
    Wscript.Echo "BlockSize: " & objItem.BlockSize
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CasePreserved: " & objItem.CasePreserved
    Wscript.Echo "CaseSensitive: " & objItem.CaseSensitive
    Wscript.Echo "CodeSet: " & objItem.CodeSet
    Wscript.Echo "CompressionMethod: " & objItem.CompressionMethod
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "EncryptionMethod: " & objItem.EncryptionMethod
    Wscript.Echo "FileSystemSize: " & objItem.FileSystemSize
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "MaxFileNameLength: " & objItem.MaxFileNameLength
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ReadOnly: " & objItem.ReadOnly
    Wscript.Echo "Root: " & objItem.Root
    Wscript.Echo "Status: " & objItem.Status
Next

