' Description: Displays information about Services for UNIX domains.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from SFU_Domain")

For Each objItem in colItems
    Wscript.Echo "DC: " & objItem.DC
    Wscript.Echo "Default: " & objItem.Default
    Wscript.Echo "Domain: " & objItem.Domain
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Maps: " & objItem.Maps
    Wscript.Echo
Next

