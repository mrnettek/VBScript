' Description: Demonstration script that modifies an IISFilterSetting property (FilterEnabled).


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFilterSetting")

For Each objItem in colItems
    objItem.FilterEnabled = TRUE
    objItem.Put_
Next

