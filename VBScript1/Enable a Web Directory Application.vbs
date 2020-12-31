' Description: Enables the application found in the Web directoryW3SVC/2142295254/root/aspnet_client_folder.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebDirectory Where Name = " & _
        "'W3SVC/2142295254/root/aspnet_client_folder'")

For Each objItem in colItems
   objItem.AppEnable(True)
Next

