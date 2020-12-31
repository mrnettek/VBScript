' Description: Continues a paused Web server named W3SVC/2142295254.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebServer Where Name = " & _
        "'W3SVC/2142295254'")

For Each objItem in colItems
    objItem.Continue
Next

