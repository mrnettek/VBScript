' Description: Displays Services for UNIX NFS server settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSServer_SvSet")

For Each objItem in colItems
    Wscript.Echo "Case Sensitive: " & objItem.CaseSensitive
    Wscript.Echo "CDFS Case: " & objItem.CdfsCase
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Directory Cache Pages: " & _
        objItem.DirectoryCachePages
    Wscript.Echo "Dot Files Hidden: " & objItem.DotFilesHidden
    Wscript.Echo "FAT Case: " & objItem.FatCase
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Logon TimeOut: " & objItem.LogonTimeOut
    Wscript.Echo "Maximum Handle Cache Size: " & _
        objItem.MaxHandleCacheSize
    Wscript.Echo "NTFS Case: " & objItem.NtfsCase
    Wscript.Echo "RdWr Handle LifeTime: " & _
        objItem.RdWrHandleLifeTime
    Wscript.Echo "Register TCP: " & objItem.RegisterTcp
    Wscript.Echo "Register Version 3: " & _
        objItem.RegisterVersion3
    Wscript.Echo
Next

