' Description: Demonstration script that displays all instances of the IIsCertMapper class.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsCertMapper")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

