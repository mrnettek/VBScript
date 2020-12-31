' Description: Displays information about Services for UNIX components.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from SFU_Component")

For Each objItem in colItems
    Wscript.Echo "Component: " & objItem.Component
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Tab Number: " & objItem.TabNum
    Wscript.Echo
Next

