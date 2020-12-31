' Description: Displays global metabase values for IIS Web site home directories.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting") 

For Each objItem in colItems
    Wscript.Echo "Application Pool ID: " & objItem.AppPoolId
    Wscript.Echo "Content Indexed: " & objItem.ContentIndexed
    Wscript.Echo "Don't Log: " & objItem.DontLog
Next

