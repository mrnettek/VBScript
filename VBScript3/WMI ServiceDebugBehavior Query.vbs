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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ServiceDebugBehavior", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "HttpHelpPageBinding: " & objItem.HttpHelpPageBinding
      WScript.Echo "HttpHelpPageEnabled: " & objItem.HttpHelpPageEnabled
      WScript.Echo "HttpHelpPageUrl: " & objItem.HttpHelpPageUrl
      WScript.Echo "HttpsHelpPageBinding: " & objItem.HttpsHelpPageBinding
      WScript.Echo "HttpsHelpPageEnabled: " & objItem.HttpsHelpPageEnabled
      WScript.Echo "HttpsHelpPageUrl: " & objItem.HttpsHelpPageUrl
      WScript.Echo "IncludeExceptionDetailInFaults: " & objItem.IncludeExceptionDetailInFaults
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

