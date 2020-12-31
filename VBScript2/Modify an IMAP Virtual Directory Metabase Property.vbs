' Description: Demonstration script that changes the AccessSource global metabase property for IMAP virtual directories on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsIMAPVirtualDirSetting")
 
For Each objItem in colItems
    objItem.AccessSource = FALSE
    objItem.Put_
Next

