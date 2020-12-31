' Description: Demonstration script that modifies IIS application configuration debugging options.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    objItem.AppAllowClientDebug = True
    objItem.AppAllowDebugging = True
    objItem.AspScriptErrorMessage = "Sorry, an error has occurred."
    objItem.AspScriptErrorSentToBrowser = True
    objItem.Put_
Next

