' Description: Lists all the applications dependant on the ASP group ID.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * From IIsWebService")
 
For Each objItem in colItems
    objItem.QueryGroupIDStatus "ASP", arrGroups
    For i = 0 to Ubound(arrGroups)
        Wscript.Echo arrGroups(i)
    Next
Next

