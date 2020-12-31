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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MicrosoftIE_Certificate", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "IssuedBy: " & objItem.IssuedBy
      WScript.Echo "IssuedTo: " & objItem.IssuedTo
      WScript.Echo "SettingID: " & objItem.SettingID
      WScript.Echo "SignatureAlgorithm: " & objItem.SignatureAlgorithm
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo "Validity: " & objItem.Validity
      WScript.Echo
   Next
Next

