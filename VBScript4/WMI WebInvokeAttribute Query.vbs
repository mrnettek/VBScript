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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM WebInvokeAttribute", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "BodyStyle: " & objItem.BodyStyle
      WScript.Echo "IsBodyStyleSetExplicitly: " & objItem.IsBodyStyleSetExplicitly
      WScript.Echo "IsRequestFormatSetExplicitly: " & objItem.IsRequestFormatSetExplicitly
      WScript.Echo "IsResponseFormatSetExplicitly: " & objItem.IsResponseFormatSetExplicitly
      WScript.Echo "Method: " & objItem.Method
      WScript.Echo "RequestFormat: " & objItem.RequestFormat
      WScript.Echo "ResponseFormat: " & objItem.ResponseFormat
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "UriTemplate: " & objItem.UriTemplate
      WScript.Echo
   Next
Next

