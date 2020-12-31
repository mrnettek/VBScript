' Description: Demonstration script that modifies a Services for UNIX NFS server auditing setting.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from NFSServer_Auditing Where KeyName = 'Parameters'")

For Each objItem in colItems
    objItem.Audit = 1
    objItem.Put_
Next

