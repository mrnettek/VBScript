' Description: List information about Web service extension restrictions.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    For i = 0 to Ubound(objItem.WebSvcExtRestrictionList)
        Wscript.Echo "Access: " & _
            objItem.WebSvcExtRestrictionList(i).Access
        Wscript.Echo "File Path: " & _
            objItem.WebSvcExtRestrictionList(i).FilePath
        Wscript.Echo
    Next
Next

