' Description: Displays Services for UNIX dummy information.


On Error Resume Next

strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from SFU_Dummy")

For Each objItem in colItems
    Wscript.Echo "Dummy: " & objItem.Dummy
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo
Next

