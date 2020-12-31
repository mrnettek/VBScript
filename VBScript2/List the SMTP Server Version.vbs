' Description: Returns the version of the SMTP Server service running on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsSmtpInfo")

For Each objItem in colItems
    Wscript.Echo "Major IIS Version Number: " & _
        objItem.MajorIIsVersionNumber
    Wscript.Echo "Minor IIS Version Number: " & _
        objItem.MinorIIsVersionNumber
    Wscript.Echo "Name: " & objItem.Name
Next

