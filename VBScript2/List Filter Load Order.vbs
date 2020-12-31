' Description: Lists filter names and load orders for all the IIS filters on a computer.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFiltersSetting")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Filter Load Order: " & objItem.FilterLoadOrder
Next

