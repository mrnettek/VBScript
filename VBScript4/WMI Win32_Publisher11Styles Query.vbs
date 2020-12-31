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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Publisher11Styles", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AutoFormatAsYouTypeDefineStyles: " & objItem.AutoFormatAsYouTypeDefineStyles
      WScript.Echo "AutoFormatPreserveStyles: " & objItem.AutoFormatPreserveStyles
      WScript.Echo "CharacterStyleCount: " & objItem.CharacterStyleCount
      WScript.Echo "ParagraphStyleCount: " & objItem.ParagraphStyleCount
      WScript.Echo
   Next
Next

