' Description: Lists all the applications in the MSSharePointAppPool application pool.


strComputer = "."

Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsApplicationPool Where Name = " & _
        "'W3SVC/AppPools/MSSharePointAppPool'")

For Each objItem in colItems
    objItem.EnumAppsInPool arrApplications
    For i = 0 to Ubound(arrApplications)
        Wscript.Echo arrApplications(i)
    Next
Next

