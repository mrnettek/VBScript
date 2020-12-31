' Description: Demonstration scripts that lists all instances of the IIsSMTPSessionsSettings class.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpSessionsSetting")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

