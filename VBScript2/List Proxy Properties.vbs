' Description: Displays proxy settings for Software Update Services.


Set objProxy = CreateObject("Microsoft.Update.WebProxy")

Wscript.Echo "Address: " & objProxy.Address
Wscript.Echo "Bypass proxy on local addresses: " & objProxy.BypassProxyOnLocal
Wscript.Echo "Read-only: " & objProxy.Readonly
Wscript.Echo "User name: " & objProxy.UserName

