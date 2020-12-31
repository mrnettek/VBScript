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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PowerPoint11PageNumber", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "PageNumbers: " & objItem.PageNumbers
      WScript.Echo "PageNumbersIn: " & objItem.PageNumbersIn
      WScript.Echo "Restart: " & objItem.Restart
      WScript.Echo "Section: " & objItem.Section
      WScript.Echo "ShowFirst: " & objItem.ShowFirst
      WScript.Echo "Start: " & objItem.Start
      WScript.Echo
   Next
Next

