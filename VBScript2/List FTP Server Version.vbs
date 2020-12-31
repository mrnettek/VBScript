' Description: Returns the version number of the FTP service running on a computer.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsFtpInfo")

For Each objItem in colItems
    Wscript.Echo "Major IIS Version Number: " & objItem.MajorIIsVersionNumber
    Wscript.Echo "Minor IIS Version Number: " & objItem.MinorIIsVersionNumber
    Wscript.Echo "Name: " & objItem.Name
Next

