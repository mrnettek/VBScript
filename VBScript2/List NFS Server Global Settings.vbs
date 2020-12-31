' Description: Displays Services for UNIX NFS server global settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSServer_Reg")

For Each objItem in colItems
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Read Configuration: " & objItem.ReadConfig
    Wscript.Echo
Next

