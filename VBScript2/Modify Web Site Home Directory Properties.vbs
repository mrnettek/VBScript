' Description: Demonstration script that modifies default Web site home directory metabase properties on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting") 

For Each objItem in colItems
    objItem.ContentIndexed = True
    objItem.DontLog = True
    objItem.Put_
Next

