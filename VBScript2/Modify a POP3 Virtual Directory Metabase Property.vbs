' Description: Demonstration script that modifies a global POP3 virtual directory metabase property value (AccessSSLRequireCert) on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsPop3VirtualDirSetting")
 
For Each objItem in colItems
    objItem.AccessSSLRequireCert = TRUE
    objItem.Put_
Next

