' Description: Displays Services for UNIX NFS server file settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSServer_Files")

For Each objItem in colItems
    Wscript.Echo "Case: " & objItem.Case
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Grace Period: " & objItem.GracePeriod
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Timeout: " & objItem.Timeout
    Wscript.Echo
Next

