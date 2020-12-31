On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_LogicalFile",,48)
For Each objItem in colItems
    Wscript.Echo "AccessMask: " & objItem.AccessMask
    Wscript.Echo "Archive: " & objItem.Archive
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Compressed: " & objItem.Compressed
    Wscript.Echo "CompressionMethod: " & objItem.CompressionMethod
    Wscript.Echo "CreationClassName: " & objItem.CreationClassName
    Wscript.Echo "CreationDate: " & objItem.CreationDate
    Wscript.Echo "CSCreationClassName: " & objItem.CSCreationClassName
    Wscript.Echo "CSName: " & objItem.CSName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Drive: " & objItem.Drive
    Wscript.Echo "EightDotThreeFileName: " & objItem.EightDotThreeFileName
    Wscript.Echo "Encrypted: " & objItem.Encrypted
    Wscript.Echo "EncryptionMethod: " & objItem.EncryptionMethod
    Wscript.Echo "Extension: " & objItem.Extension
    Wscript.Echo "FileName: " & objItem.FileName
    Wscript.Echo "FileSize: " & objItem.FileSize
    Wscript.Echo "FileType: " & objItem.FileType
    Wscript.Echo "FSCreationClassName: " & objItem.FSCreationClassName
    Wscript.Echo "FSName: " & objItem.FSName
    Wscript.Echo "Hidden: " & objItem.Hidden
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "InUseCount: " & objItem.InUseCount
    Wscript.Echo "LastAccessed: " & objItem.LastAccessed
    Wscript.Echo "LastModified: " & objItem.LastModified
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo "Readable: " & objItem.Readable
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "System: " & objItem.System
    Wscript.Echo "Writeable: " & objItem.Writeable
Next

