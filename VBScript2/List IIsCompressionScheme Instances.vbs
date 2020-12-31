' Description: Demonstration script that displays all instances of the IIsCompressionScheme class.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsCompressionScheme")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

