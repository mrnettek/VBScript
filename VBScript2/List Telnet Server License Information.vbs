' Description: Displays Services for UNIX Telnet server license information.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from TelnetServer_Licences")

For Each objItem in colItems
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Licenses: " & objItem.Licences
    Wscript.Echo "Mode: " & objItem.Mode
    Wscript.Echo
Next

