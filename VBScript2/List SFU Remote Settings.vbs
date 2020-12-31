' Description: Displays Services for UNIX remote settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from SFU_Remote")

For Each objItem in colItems
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Machine: " & objItem.Machine
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo
Next

