' Description: Lists the name of the computer where IIS is running.


strComputer = "."
 
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsComputer")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

