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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM SymmetricSecurityBindingElement", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "DefaultAlgorithmSuite: " & objItem.DefaultAlgorithmSuite
      WScript.Echo "IncludeTimestamp: " & objItem.IncludeTimestamp
      WScript.Echo "KeyEntropyMode: " & objItem.KeyEntropyMode
      WScript.Echo "LocalServiceSecuritySettings: " & objItem.LocalServiceSecuritySettings
      WScript.Echo "MessageProtectionOrder: " & objItem.MessageProtectionOrder
      WScript.Echo "MessageSecurityVersion: " & objItem.MessageSecurityVersion
      WScript.Echo "RequireSignatureConfirmation: " & objItem.RequireSignatureConfirmation
      WScript.Echo "SecurityHeaderLayout: " & objItem.SecurityHeaderLayout
      WScript.Echo
   Next
Next

