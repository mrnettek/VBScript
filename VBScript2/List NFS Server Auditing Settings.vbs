' Description: Displays Services for UNIX NFS server auditing settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSServer_Auditing")

For Each objItem in colItems
    Wscript.Echo "Audit: " & objItem.Audit
    Wscript.Echo "Audit Bits: " & objItem.AuditBits
    Wscript.Echo "Check Space: " & objItem.CheckSpace
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "File Maximum Size: " & objItem.FileMaxSize
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Log File: " & objItem.LogFile
    Wscript.Echo "Minimum Space: " & objItem.MinSpace
    Wscript.Echo
Next

