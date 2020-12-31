' Description: Demonstration script that modifies IIS application pools recycling settings.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsApplicationPoolsSetting")

For Each objItem in colItems
    objItem.PeriodicRestartMemory = 1000000
    objItem.PeriodicRestartPrivateMemory = 1000000
    objItem.PeriodicRestartRequests = 5
    objItem.PeriodicRestartTime = 3480
    objItem.Put_
Next

