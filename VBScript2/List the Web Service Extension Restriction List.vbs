' Description: Lists all items on the Web service extension restriction list.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    For i = 0 to Ubound(objItem.WebSvcExtRestrictionList)
        Wscript.Echo "Server Extension: " & _  
            objItem.WebSvcExtRestrictionList(i)._
                ServerExtension
        Wscript.Echo "Access: " & _
            objItem.WebSvcExtRestrictionList(i).Access
        Wscript.Echo "Deletable: " & _
            objItem.WebSvcExtRestrictionList(i).Deletable
        Wscript.Echo "Description: " & _
            objItem.WebSvcExtRestrictionList(i).Description
        Wscript.Echo "File Path: " & _
            objItem.WebSvcExtRestrictionList(i).FilePath
        Wscript.Echo
    Next
Next

