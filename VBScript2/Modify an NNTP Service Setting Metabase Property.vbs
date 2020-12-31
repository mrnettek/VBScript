' Description: Demonstration script that modifies a global metabase property (ContentIndexed) for the NNTP server service on a computer.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNNTPServiceSetting")
 
For Each objItem in colItems
    objItem.ContentIndexed = TRUE
    objItem.Put_
Next

