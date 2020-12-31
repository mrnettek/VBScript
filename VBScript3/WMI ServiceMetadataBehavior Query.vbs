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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM ServiceMetadataBehavior", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ExternalMetadataLocation: " & objItem.ExternalMetadataLocation
      WScript.Echo "HttpGetBinding: " & objItem.HttpGetBinding
      WScript.Echo "HttpGetEnabled: " & objItem.HttpGetEnabled
      WScript.Echo "HttpGetUrl: " & objItem.HttpGetUrl
      WScript.Echo "HttpsGetBinding: " & objItem.HttpsGetBinding
      WScript.Echo "HttpsGetEnabled: " & objItem.HttpsGetEnabled
      WScript.Echo "HttpsGetUrl: " & objItem.HttpsGetUrl
      WScript.Echo "MetadataExportInfo: " & objItem.MetadataExportInfo
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

