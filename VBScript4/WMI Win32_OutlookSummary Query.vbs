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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OutlookSummary", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Build: " & objItem.Build
      WScript.Echo "Folder: " & objItem.Folder
      WScript.Echo "Item: " & objItem.Item
      WScript.Echo "Language: " & objItem.Language
      WScript.Echo "MailSupport: " & objItem.MailSupport
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Path: " & objItem.Path
      WScript.Echo "ProductID: " & objItem.ProductID
      WScript.Echo "SystemLanguage: " & objItem.SystemLanguage
      WScript.Echo "User: " & objItem.User
      WScript.Echo "Version: " & objItem.Version
      WScript.Echo
   Next
Next

