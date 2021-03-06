' Description: Lists IIS application configuration cache options.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    Wscript.Echo "ASP Disk Template Cache Directory: " & _
        objItem.AspDiskTemplateCacheDirectory
    Wscript.Echo "ASP Maximum Disk Template Cache Files: " & _
        objItem.AspMaxDiskTemplateCacheFiles
    Wscript.Echo "ASP Script Engine Cache Maximum: " & _
        objItem.AspScriptEngineCacheMax
    Wscript.Echo "ASP Script File Cache Size: " & _
        objItem.AspScriptFileCacheSize
Next

