' Description: Demonstration script that changes the DefaultLogonDomain value for all the IMAP servers on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsIMAPServerSetting")
 
For Each objItem in colItems
    objItem.DefaultLogonDomain = "fabrikam"
    objItem.Put_
Next

