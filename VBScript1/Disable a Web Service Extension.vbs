' Description: Disables a Web service extension named WEBDAV.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * From IIsWebService")

For Each objItem in colItems
    objItem.DisableWebServiceExtension("WEBDAV")
Next

