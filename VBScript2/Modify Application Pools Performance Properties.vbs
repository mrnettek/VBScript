' Description: Demonstration script that modifies performance settings for IIS application pools.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsApplicationPoolsSetting")

For Each objItem in colItems
    objItem.AppPoolQueueLength = 5000
    objItem.CPUAction = 1
    objItem.CPULimit = 50000
    objItem.CPUResetInterval= 30
    objItem.IdleTimeout = 30
    objItem.MaxProcesses = 2
    objItem.Put_
Next

