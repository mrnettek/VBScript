' Description: Demonstration script that displays all instances of the IIsFilter class.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsFilter")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

