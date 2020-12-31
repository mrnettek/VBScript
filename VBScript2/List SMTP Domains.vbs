' Description: Lists the names of all the SMTP domains on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsSmtpDomain")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

