' Description: Demonstration script that modifies the VrDoExpire global NNTP metabase property on a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNNTPVirtualDirSetting")
 
For Each objItem in colItems
    objItem.VrDoExpire = TRUE
    objItem.Put_
Next

