' Description: Displays the Services for UNIX Windows directory.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from SFU_Windir")

For Each objItem in colItems
    Wscript.Echo "Default: " & objItem.default
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Windows directory: " & objItem.windir
    Wscript.Echo
Next

