' Description: Lists all instances of the IIsObjectSetting class.


strComputer = "."
 
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsObjectSetting")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

