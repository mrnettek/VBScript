On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_NFS",,48)
For Each objItem in colItems
    Wscript.Echo "AttributeCaching: " & objItem.AttributeCaching
    Wscript.Echo "AttributeCachingForDirectoriesMax: " & objItem.AttributeCachingForDirectoriesMax
    Wscript.Echo "AttributeCachingForDirectoriesMin: " & objItem.AttributeCachingForDirectoriesMin
    Wscript.Echo "AttributeCachingForRegularFilesMax: " & objItem.AttributeCachingForRegularFilesMax
    Wscript.Echo "AttributeCachingForRegularFilesMin: " & objItem.AttributeCachingForRegularFilesMin
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
    Wscript.Echo "ForegroundMount: " & objItem.ForegroundMount
    Wscript.Echo "HardMount: " & objItem.HardMount
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Interrupt: " & objItem.Interrupt
    Wscript.Echo "MaxFileNameLength: " & objItem.MaxFileNameLength
    Wscript.Echo "MountFailureRetries: " & objItem.MountFailureRetries
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "ReadBufferSize: " & objItem.ReadBufferSize
    Wscript.Echo "ReadOnly: " & objItem.ReadOnly
    Wscript.Echo "RetransmissionAttempts: " & objItem.RetransmissionAttempts
    Wscript.Echo "RetransmissionTimeout: " & objItem.RetransmissionTimeout
    Wscript.Echo "Root: " & objItem.Root
    Wscript.Echo "ServerCommunicationPort: " & objItem.ServerCommunicationPort
    Wscript.Echo "Status: " & objItem.Status
    Wscript.Echo "WriteBufferSize: " & objItem.WriteBufferSize
Next

