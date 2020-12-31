' Description: Displays mapping server settings for Services for UNIX.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from MapServer_Reg")

For Each objItem in colItems
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "KeyName: " & objItem.KeyName
    Wscript.Echo "ReadConfig: " & objItem.ReadConfig
    Wscript.Echo
Next

