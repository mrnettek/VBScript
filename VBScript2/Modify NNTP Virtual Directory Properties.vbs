' Description: Demonstration script that modifies global NNTP default property values on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpVirtualDirSetting")

For Each objItem in colItems
    objItem.ContentIndexed = True
    objItem.DontLog = False
    objItem.Path = "c:\inetpub\nntpfile\root\accounting"
    objItem.Put_
Next

