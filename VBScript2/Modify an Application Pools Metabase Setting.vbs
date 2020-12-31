' Description: Demonstration script that modifies an application pools metabase setting (RapidFailProtection).


strComputer = "."
 
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsApplicationPoolsSetting")
 
For Each objItem in colItems
    objItem.RapidFailProtection = TRUE
    objItem.Put_
Next

