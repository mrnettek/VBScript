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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ServiceSecurityAuditBehavior", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AuditLogLocation: " & objItem.AuditLogLocation
      WScript.Echo "MessageAuthenticationAuditLevel: " & objItem.MessageAuthenticationAuditLevel
      WScript.Echo "ServiceAuthorizationAuditLevel: " & objItem.ServiceAuthorizationAuditLevel
      WScript.Echo "SuppressAuditFailure: " & objItem.SuppressAuditFailure
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

