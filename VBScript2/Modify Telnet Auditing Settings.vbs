' Description: Demonstration script that modifies a Services for UNIX Telnet auditing setting.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Telnet_Auditing Where KeyName = 'Defaults'")

For Each objItem in colItems
    objItem.LogEvents = 1
    objItem.Put_
Next

