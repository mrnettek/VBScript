' Description: Displays a list of IIS filters.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsFilters")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

