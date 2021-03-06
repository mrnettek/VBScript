' Description: Verifies whether or not IIS is running in worker process isolation mode.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebService")

For Each objItem in colItems
    intMode = objItem.GetCurrentMode
    If intMode = 0 Then
        Wscript.Echo _
            "IIS is in IIS 5.0 isolation mode."
    ElseIf intMode = 1 Then
        Wscript.Echo _
            "IIS is in worker process isolation mode."
    Else
        Wscript.Echo _
            "The current mode cannot be determined."
    End If
Next

