' Description: Displays Services for UNIX NFS client performance settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSClient_Perf Where KeyName = 'Defaults'")

For Each objItem in colItems
    objItem.Retries = 2
    objItem.Put_
Next

