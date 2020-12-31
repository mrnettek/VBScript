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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_AccessSummary", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ADOVersion: " & objItem.ADOVersion
      WScript.Echo "Build: " & objItem.Build
      WScript.Echo "Language: " & objItem.Language
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Path: " & objItem.Path
      WScript.Echo "ProductID: " & objItem.ProductID
      WScript.Echo "References: " & objItem.References
      WScript.Echo "VBAVersion: " & objItem.VBAVersion
      WScript.Echo "Version: " & objItem.Version
      WScript.Echo
   Next
Next

