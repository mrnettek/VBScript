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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MicrosoftIE_Summary", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ActivePrinter: " & objItem.ActivePrinter
      WScript.Echo "Build: " & objItem.Build
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "CipherStrength: " & objItem.CipherStrength
      WScript.Echo "ContentAdvisor: " & objItem.ContentAdvisor
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "IEAKInstall: " & objItem.IEAKInstall
      WScript.Echo "Language: " & objItem.Language
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Path: " & objItem.Path
      WScript.Echo "ProductID: " & objItem.ProductID
      WScript.Echo "SettingID: " & objItem.SettingID
      WScript.Echo "Version: " & objItem.Version
      WScript.Echo
   Next
Next

