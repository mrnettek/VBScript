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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PowerPoint11Hyperlink", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Address: " & objItem.Address
      WScript.Echo "Index: " & objItem.Index
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "ScreenTip: " & objItem.ScreenTip
      WScript.Echo "Subaddress: " & objItem.Subaddress
      WScript.Echo "Target: " & objItem.Target
      WScript.Echo "TextToDisplay: " & objItem.TextToDisplay
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

