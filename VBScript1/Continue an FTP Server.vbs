' Description: Resumes an FTP server named MSFTPSVC/1.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsFtpServer Where Name = " & _
        "'MSFTPSVC/1'")

For Each objItem in colItems
    objItem.Continue
Next

