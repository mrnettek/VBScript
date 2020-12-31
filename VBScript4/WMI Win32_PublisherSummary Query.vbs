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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PublisherSummary", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ActivePrinter: " & objItem.ActivePrinter
      WScript.Echo "DisplayName: " & objItem.DisplayName
      WScript.Echo "DisplayVersion: " & objItem.DisplayVersion
      WScript.Echo "HelpLink: " & objItem.HelpLink
      WScript.Echo "InstallDate: " & objItem.InstallDate
      WScript.Echo "InstallLocation: " & objItem.InstallLocation
      WScript.Echo "InstallSource: " & objItem.InstallSource
      WScript.Echo "Language: " & objItem.Language
      WScript.Echo "LocalPackage: " & objItem.LocalPackage
      WScript.Echo "ProductID: " & objItem.ProductID
      WScript.Echo "RegCompany: " & objItem.RegCompany
      WScript.Echo "RegOwner: " & objItem.RegOwner
      WScript.Echo "SystemLanguage: " & objItem.SystemLanguage
      WScript.Echo "URLInfoAbout: " & objItem.URLInfoAbout
      WScript.Echo "URLUpdateInfo: " & objItem.URLUpdateInfo
      WScript.Echo
   Next
Next

