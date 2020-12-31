' Description: Returns information about the Internet proxy server used by a computer.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Proxy")

For Each objItem in colItems
    Wscript.Echo "Proxy Port Number: " & objItem.ProxyPortNumber
    Wscript.Echo "Proxy Server: " & objItem.ProxyServer
    Wscript.Echo "Server Name: " & objItem.ServerName
    Wscript.Echo
Next

