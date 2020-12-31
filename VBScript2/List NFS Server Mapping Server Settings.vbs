' Description: Displays Services for UNIX NFS server mapping server settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSServer_MapSvr")

For Each objItem in colItems
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Map Server: " & objItem.MapSvr
    Wscript.Echo
Next

