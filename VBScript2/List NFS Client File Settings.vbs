' Description: Displays Services for UNIX NFS client file settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSClient_Files")

For Each objItem in colItems
    Wscript.Echo "Access: " & objItem.Access
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo
Next

