' Description: Returns Microsoft Exchange 2003 schedule interval information.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_ScheduleInterval")

For Each objItem in colItems
    Wscript.Echo "Start time: " & objItem.StartTime
    Wscript.Echo "Stop time: " & objItem.StopTime
    Wscript.Echo
Next

