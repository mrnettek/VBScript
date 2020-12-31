' Description: Displays Services for UNIX NFS server security settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSServer_Security")

For Each objItem in colItems
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Security: " & objItem.Security
    Wscript.Echo
Next

