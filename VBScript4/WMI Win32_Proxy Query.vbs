On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Proxy",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ProxyPortNumber: " & objItem.ProxyPortNumber
    Wscript.Echo "ProxyServer: " & objItem.ProxyServer
    Wscript.Echo "ServerName: " & objItem.ServerName
    Wscript.Echo "SettingID: " & objItem.SettingID
Next

