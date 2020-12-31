On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\ServiceModel")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM HttpTransportBindingElement", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AllowCookies: " & objItem.AllowCookies
      WScript.Echo "AuthenticationScheme: " & objItem.AuthenticationScheme
      WScript.Echo "BypassProxyOnLocal: " & objItem.BypassProxyOnLocal
      WScript.Echo "HostNameComparisonMode: " & objItem.HostNameComparisonMode
      WScript.Echo "KeepAliveEnabled: " & objItem.KeepAliveEnabled
      WScript.Echo "ManualAddressing: " & objItem.ManualAddressing
      WScript.Echo "MaxBufferPoolSize: " & objItem.MaxBufferPoolSize
      WScript.Echo "MaxBufferSize: " & objItem.MaxBufferSize
      WScript.Echo "MaxReceivedMessageSize: " & objItem.MaxReceivedMessageSize
      WScript.Echo "ProxyAddress: " & objItem.ProxyAddress
      WScript.Echo "ProxyAuthenticationScheme: " & objItem.ProxyAuthenticationScheme
      WScript.Echo "Realm: " & objItem.Realm
      WScript.Echo "Scheme: " & objItem.Scheme
      WScript.Echo "TransferMode: " & objItem.TransferMode
      WScript.Echo "UnsafeConnectionNtlmAuthentication: " & objItem.UnsafeConnectionNtlmAuthentication
      WScript.Echo "UseDefaultWebProxy: " & objItem.UseDefaultWebProxy
      WScript.Echo
   Next
Next

