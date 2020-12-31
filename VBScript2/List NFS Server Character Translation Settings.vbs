' Description: Displays Services for UNIX NFS server character translation settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSServer_SvSet_CharTran")

For Each objItem in colItems
    Wscript.Echo "Character Translation: " & objItem.CharacterTranslation
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo
Next

