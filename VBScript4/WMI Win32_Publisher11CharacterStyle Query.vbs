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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Publisher11CharacterStyle", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "BaseStyle: " & objItem.BaseStyle
      WScript.Echo "BuiltIn: " & objItem.BuiltIn
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo
   Next
Next

