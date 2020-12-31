' Description: Disables the BITSsrv.dll extension file.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebService")

For Each objItem in colItems
    objItem.DisableExtensionFile _
        ("C:\WINDOWS\system32\bitssrv.dll")
Next

