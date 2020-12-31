' Description: Demonstration script that modifies the global POP3 server AuthAnonymous property metabase value on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsPop3ServiceSetting")
 
For Each objItem in colItems
    objItem.AuthAnonymous = FALSE
    objItem.Put_
Next

