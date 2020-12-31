' Description: Demonstration script that modifies the ServerAutoStart global IMAP metabase property on a computer IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery( _
    "Select * from IIsIMAPServiceSetting")
 
For Each objItem in colItems
    objItem.ServerAutoStart = TRUE
    objItem.Put_
Next

