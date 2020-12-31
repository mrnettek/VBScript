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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ServiceCredentials", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ClientCertificate: " & objItem.ClientCertificate
      WScript.Echo "IssuedTokenAuthentication: " & objItem.IssuedTokenAuthentication
      WScript.Echo "Peer: " & objItem.Peer
      WScript.Echo "SecureConversationAuthentication: " & objItem.SecureConversationAuthentication
      WScript.Echo "ServiceCertificate: " & objItem.ServiceCertificate
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "UserNameAuthentication: " & objItem.UserNameAuthentication
      WScript.Echo "WindowsAuthentication: " & objItem.WindowsAuthentication
      WScript.Echo
   Next
Next

