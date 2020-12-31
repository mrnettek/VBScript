' Description: Demonstration script that stops the MSSharePointAppPool application pool.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsApplicationPool Where Name = " & _
        "'W3SVC/AppPools/MSSharePointAppPool'")

For Each objItem in colItems
    objItem.Stop
Next

