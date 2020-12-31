' Description: Demonstration script that modifies Web service logging options on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    objItem.LogFileDirectory = "C:\Logs"
    objItem.LogFileLocaltimeRollover = True
    objItem.LogFilePeriod = 2
    objItem.LogFileTruncateSize = 1000000
    objItem.Put_
Next

