' Description: Lists IIS Web service applications.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebService")

For Each objItem in colItems
    objItem.ListApplications arrApplications
    For i = 0 to Ubound(arrApplications)
        Wscript.Echo arrApplications(i)
    Next
Next

