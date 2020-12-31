On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\MSAPPS11")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Word11Summary", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ActivePrinter: " & objItem.ActivePrinter
      WScript.Echo "AddinCount: " & objItem.AddinCount
      WScript.Echo "Build: " & objItem.Build
      WScript.Echo "DocumentCount: " & objItem.DocumentCount
      WScript.Echo "Language: " & objItem.Language
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Path: " & objItem.Path
      WScript.Echo "ProductID: " & objItem.ProductID
      WScript.Echo "SystemLanguage: " & objItem.SystemLanguage
      WScript.Echo "TemplateCount: " & objItem.TemplateCount
      WScript.Echo "Version: " & objItem.Version
      WScript.Echo
   Next
Next

