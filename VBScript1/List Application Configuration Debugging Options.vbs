' Description: Lists IIS application configuration debugging options.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    Wscript.Echo "Application Allow Client Debug: " & _
        objItem.AppAllowClientDebug
    Wscript.Echo "Application Allow Debugging: " & _
        objItem.AppAllowDebugging
   Wscript.Echo "ASP Script Error Message: " & _
        objItem.AspScriptErrorMessage
    Wscript.Echo "ASP Script Error Sent To Browser: " & _
        objItem.AspScriptErrorSentToBrowser
Next

