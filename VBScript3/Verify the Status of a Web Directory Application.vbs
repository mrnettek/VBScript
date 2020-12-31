' Description: Returns the status of the application in the W3SVC/2142295254/root/aspnet_client_folder Web directory.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebDirectory Where Name = " & _
        "'W3SVC/2142295254/root/aspnet_client_folder'")

For Each objItem in colItems
    strStatus = objItem.AppGetStatus
    If strStatus = 2 Then
        Wscript.Echo "The application is running."
    ElseIf strStatus = 3 Then
        Wscript.Echo "The application is stopped."
    Else
        Wscript.Echo _
            "The status could not be determined."
    End If
Next

