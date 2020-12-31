' Description: Demonstration script that modifies a Services for UNIX Telnet server setting.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _\
    ("Select * from TelnetServer_Settings Where KeyName = 'Defaults'")

For Each objItem in colItems
    objItem.DefaultDomain = "fabrikam.com"
    objItem.Put_
Next

