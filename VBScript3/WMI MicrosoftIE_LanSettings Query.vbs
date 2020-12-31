On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2\Applications\MicrosoftIE")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MicrosoftIE_LanSettings", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AutoConfigProxy: " & objItem.AutoConfigProxy
      WScript.Echo "AutoConfigURL: " & objItem.AutoConfigURL
      WScript.Echo "AutoProxyDetectMode: " & objItem.AutoProxyDetectMode
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "Proxy: " & objItem.Proxy
      WScript.Echo "ProxyOverride: " & objItem.ProxyOverride
      WScript.Echo "ProxyServer: " & objItem.ProxyServer
      WScript.Echo "SettingID: " & objItem.SettingID
      WScript.Echo
   Next
Next

