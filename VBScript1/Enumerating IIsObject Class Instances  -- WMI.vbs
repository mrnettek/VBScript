' Description: Lists all the instances of the IIsObject class.


strComputer = "."
 
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsObject")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

