' Description: Displays Services for UNIX mapper settings.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\sfuadmin")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Mapper_Settings Where KeyName = 'CurrentVersion'")

For Each objItem in colItems
    objItem.LoggingLevel = 1
    objItem.Put_
Next

