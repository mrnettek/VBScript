' Description: Lists IIS application configuration mappings.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    Wscript.Echo "Cache ISAPI: " & objItem.CacheISAPI
    For i = 0 to Ubound(objItem.ScriptMaps)
        Wscript.Echo "Extension: " & objItem.ScriptMaps(i).Extensions
        Wscript.Echo "Included Verbs: " & _
            objItem.ScriptMaps(i).IncludedVerbs
        Wscript.Echo "Script Processor: " & _
            objItem.ScriptMaps(i).ScriptProcessor
        Wscript.Echo
    Next
Next

