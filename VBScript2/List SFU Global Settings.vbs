' Description: Displays global Services for UNIX settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from GlobalSettings_Reg")

For Each objItem in colItems
    Wscript.Echo "Dummy: " & objItem.Dummy
    Wscript.Echo "Key Name: " & objItem.KeyName
    For Each strServer in objItem.MappingServers      
        Wscript.Echo "Mapping Server: " & strServer
    Next
    Wscript.Echo
Next

