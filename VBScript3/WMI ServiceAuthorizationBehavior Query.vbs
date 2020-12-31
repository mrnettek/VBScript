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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ServiceAuthorizationBehavior", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ImpersonateCallerForAllOperations: " & objItem.ImpersonateCallerForAllOperations
      WScript.Echo "PrincipalPermissionMode: " & objItem.PrincipalPermissionMode
      WScript.Echo "RoleProvider: " & objItem.RoleProvider
      WScript.Echo "ServiceAuthorizationManager: " & objItem.ServiceAuthorizationManager
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

