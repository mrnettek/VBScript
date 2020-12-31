' Description: Demonstration script that modifies IIS application pools health properties.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsApplicationPoolsSetting")

For Each objItem in colItems
    objItem.PingingEnabled = True
    objItem.PingInterval = 60
    objItem.RapidFailProtection = True
    objItem.RapidFailProtectionInterval = 10
    objItem.RapidFailProtectionMaxCrashes = 10
    objItem.ShutdownTimeLimit = 120
    objItem.StartupTimeLimit = 120
    objItem.Put_
Next

