' Description: Displays Services for UNIX NFS client performance information.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSClient_Perf")

For Each objItem in colItems
    Wscript.Echo "AutoTuning: " & objItem.AutoTuning
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Defaults: " & objItem.Defaults
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Mount Type: " & objItem.MountType
    Wscript.Echo "Prefer TCP: " & objItem.PreferTCP
    Wscript.Echo "Read Buffer: " & objItem.ReadBuffer
    Wscript.Echo "Retries: " & objItem.Retries
    Wscript.Echo "Timeout: " & objItem.Timeout
    Wscript.Echo "Write Buffer: " & objItem.WriteBuffer
    Wscript.Echo
Next

