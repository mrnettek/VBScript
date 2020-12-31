' Description: Creates an out-of-process application in the Web directory W3SVC/2142295254/root/aspnet_client_folder.


Const OUT_OF_PROCESS = 1

Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebDirectory Where Name = " & _
        "'W3SVC/2142295254/root/aspnet_client_folder'")

For Each objItem in colItems
   objItem.AppCreate2(OUT_OF_PROCESS)
Next

