' Description: Demonstration script that modifies IIS local computer properties.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsComputerSetting")

For Each objItem in colItems
    objItem.EnableEditWhileRunning = 1
    objItem.EnableHistory = 1
    objItem.MaxHistoryFiles = 50
    objItem.Put_
Next

