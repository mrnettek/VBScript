' Description: Displays Services for UNIX Telnet server authentication settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Telnet_Authenticate")

For Each objItem in colItems
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Mechanism: " & objItem.Mechanism
    Wscript.Echo
Next

