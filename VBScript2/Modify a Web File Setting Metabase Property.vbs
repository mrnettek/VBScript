' Description: Demonstration script that modifies a global Web file metabase property (EnableDocFooter) on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebFileSetting")

For Each objItem in colItems
    objItem.EnableDocFooter = TRUE
    objItem.Put_
Next

