' Description: Demonstration script that modifies the AuthPassport property value for all the POIP3 servers on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsPop3ServerSetting")
 
For Each objItem in colItems
    objItem.AuthPassport = TRUE
    objItem.Put_
Next

