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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ClientCredentials", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ClientCertificate: " & objItem.ClientCertificate
      WScript.Echo "HttpDigest: " & objItem.HttpDigest
      WScript.Echo "IssuedToken: " & objItem.IssuedToken
      WScript.Echo "Peer: " & objItem.Peer
      WScript.Echo "ServiceCertificate: " & objItem.ServiceCertificate
      WScript.Echo "SupportInteractive: " & objItem.SupportInteractive
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "UserName: " & objItem.UserName
      WScript.Echo "Windows: " & objItem.Windows
      WScript.Echo
   Next
Next

