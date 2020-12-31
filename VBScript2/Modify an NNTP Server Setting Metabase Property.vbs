' Description: Demonstration script that modifies a metabase property (AllowClientPosts) for all NNTP server sites on a computer.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNNTPServerSetting")
 
For Each objItem in colItems
    objItem.AllowClientPosts = TRUE
    objItem.Put_
Next

