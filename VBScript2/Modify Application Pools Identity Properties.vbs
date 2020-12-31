' Description: Demonstration script that modifies the user name and password for IIS application pools.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsApplicationPoolsSetting")

For Each objItem in colItems
    objItem.WAMUserName = "TestUser"
    objItem.WAMUserPass = "ur^354Hdf"
    objItem.Put_
Next

