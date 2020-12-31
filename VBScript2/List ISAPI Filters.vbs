' Description: Displays a list of ISAPI filters, their path, and their current state.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFilterSetting")

For Each objItem in colItems
    Wscript.Echo "Filter Path: " & objItem.FilterPath
    Wscript.Echo "Filter State: " & objItem.FilterState
    Wscript.Echo "Name: " & objItem.Name
Next

