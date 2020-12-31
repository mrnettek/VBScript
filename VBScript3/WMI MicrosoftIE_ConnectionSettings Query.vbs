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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MicrosoftIE_ConnectionSettings", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AllowInternetPrograms: " & objItem.AllowInternetPrograms
      WScript.Echo "AutoConfigURL: " & objItem.AutoConfigURL
      WScript.Echo "AutoDisconnect: " & objItem.AutoDisconnect
      WScript.Echo "AutoProxyDetectMode: " & objItem.AutoProxyDetectMode
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "DataEncryption: " & objItem.DataEncryption
      WScript.Echo "Default: " & objItem.Default
      WScript.Echo "DefaultGateway: " & objItem.DefaultGateway
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "DialUpServer: " & objItem.DialUpServer
      WScript.Echo "DisconnectIdleTime: " & objItem.DisconnectIdleTime
      WScript.Echo "EncryptedPassword: " & objItem.EncryptedPassword
      WScript.Echo "IPAddress: " & objItem.IPAddress
      WScript.Echo "IPHeaderCompression: " & objItem.IPHeaderCompression
      WScript.Echo "Modem: " & objItem.Modem
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "NetworkLogon: " & objItem.NetworkLogon
      WScript.Echo "NetworkProtocols: " & objItem.NetworkProtocols
      WScript.Echo "PrimaryDNS: " & objItem.PrimaryDNS
      WScript.Echo "PrimaryWINS: " & objItem.PrimaryWINS
      WScript.Echo "Proxy: " & objItem.Proxy
      WScript.Echo "ProxyOverride: " & objItem.ProxyOverride
      WScript.Echo "ProxyServer: " & objItem.ProxyServer
      WScript.Echo "RedialAttempts: " & objItem.RedialAttempts
      WScript.Echo "RedialWait: " & objItem.RedialWait
      WScript.Echo "ScriptFileName: " & objItem.ScriptFileName
      WScript.Echo "SecondaryDNS: " & objItem.SecondaryDNS
      WScript.Echo "SecondaryWINS: " & objItem.SecondaryWINS
      WScript.Echo "ServerAssignedIPAddress: " & objItem.ServerAssignedIPAddress
      WScript.Echo "ServerAssignedNameServer: " & objItem.ServerAssignedNameServer
      WScript.Echo "SettingID: " & objItem.SettingID
      WScript.Echo "SoftwareCompression: " & objItem.SoftwareCompression
      WScript.Echo
   Next
Next

