' Description: Lists the properties for the file C:\Scripts\Adsi.vbs.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_Datafile Where name = 'c:\\Scripts\\Adsi.vbs'")

For Each objFile in colFiles
    Wscript.Echo "Access mask: " & objFile.AccessMask
    Wscript.Echo "Archive: " & objFile.Archive
    Wscript.Echo "Compressed: " & objFile.Compressed
    Wscript.Echo "Compression method: " & objFile.CompressionMethod
    Wscript.Echo "Creation date: " & objFile.CreationDate
    Wscript.Echo "Computer system name: " & objFile.CSName
    Wscript.Echo "Drive: " & objFile.Drive
    Wscript.Echo "8.3 file name: " & objFile.EightDotThreeFileName
    Wscript.Echo "Encrypted: " & objFile.Encrypted
    Wscript.Echo "Encryption method: " & objFile.EncryptionMethod
    Wscript.Echo "Extension: " & objFile.Extension
    Wscript.Echo "File name: " & objFile.FileName
    Wscript.Echo "File size: " & objFile.FileSize
    Wscript.Echo "File type: " & objFile.FileType
    Wscript.Echo "File system name: " & objFile.FSName
    Wscript.Echo "Hidden: " & objFile.Hidden
    Wscript.Echo "Last accessed: " & objFile.LastAccessed
    Wscript.Echo "Last modified: " & objFile.LastModified
    Wscript.Echo "Manufacturer: " & objFile.Manufacturer
    Wscript.Echo "Name: " & objFile.Name
    Wscript.Echo "Path: " & objFile.Path
    Wscript.Echo "Readable: " & objFile.Readable
    Wscript.Echo "System: " & objFile.System
    Wscript.Echo "Version: " & objFile.Version
    Wscript.Echo "Writeable: " & objFile.Writeable
Next

