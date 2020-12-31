' Description: Demonstration script that modifies IIS application configuration cache options.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    objItem.AspDiskTemplateCacheDirectory = "C:\Cache"
    objItem.AspMaxDiskTemplateCacheFiles = 5000
    objItem.AspScriptEngineCacheMax = 250
    objItem.AspScriptFileCacheSize = 500
    objItem.Put_
Next

