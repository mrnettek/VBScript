' Description: Lists all IIS Web service extensions.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebService")

For Each objItem in colItems
    objItem.ListWebServiceExtensions arrExtensions
    For i = 0 to Ubound(arrExtensions)
        Wscript.Echo arrExtensions(i)
    Next
Next

