' Description: Returns the names of all the SMTP sessions running on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsSmtpSessions")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

