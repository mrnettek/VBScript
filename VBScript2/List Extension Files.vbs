' Description: Lists all IIS extension files.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebService")

For Each objItem in colItems
    objItem.ListExtensionFiles arrFiles
    For i = 0 to Ubound(arrFiles)
        Wscript.Echo arrFiles(i)
    Next
Next

