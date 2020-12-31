' Description: Lists IIS server binding information.


strComputer = "."

Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServerSetting")

For Each objItem in colItems
    For i = 0 to Ubound(objItem.ServerBindings)
        Wscript.Echo "Port: " & _
            objItem.ServerBindings(i).Port
    Next
Next

