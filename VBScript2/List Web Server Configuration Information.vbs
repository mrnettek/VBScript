' Description: Returns global Web info metabase property values for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsWebInfoSetting")
 
For Each objItem in colItems
    Wscript.Echo "Admin Server: " & objItem.AdminServer
    Wscript.Echo "Log Module List: " & objItem.LogModuleList
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Server Configuration Auto Password Sync: " & _
        objItem.ServerConfigAutoPWSync
    Wscript.Echo "Server Configuration Flags: " & _
        objItem.ServerConfigFlags
    Wscript.Echo "Server Configuration SSL 128: " & _
        objItem.ServerConfigSSL128
    Wscript.Echo "Server Configuration SSL 40: " & _
        objItem.ServerConfigSSL40
    Wscript.Echo "Server Configuration SSL Allow Encryption: " & _
        objItem.ServerConfigSSLAllowEncrypt
    Wscript.Echo "Setting ID: " & objItem.SettingID
Next

