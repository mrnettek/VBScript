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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Publisher11DefaultFileLocation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Path: " & objItem.Path
      WScript.Echo
   Next
Next

