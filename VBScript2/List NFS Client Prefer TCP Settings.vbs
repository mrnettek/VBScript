' Description: Displays Services for UNIX NFS client TCP preference information.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSClient_PreferTCP")

For Each objItem in colItems
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Prefer TCP: " & objItem.PreferTCP
    Wscript.Echo
Next

