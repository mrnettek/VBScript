' Description: Lists IIS MIME maps.


strComputer = "."
 
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsMimeMap")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

