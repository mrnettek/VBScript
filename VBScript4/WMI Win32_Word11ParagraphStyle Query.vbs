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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Word11ParagraphStyle", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AutomaticallyUpdate: " & objItem.AutomaticallyUpdate
      WScript.Echo "BaseStyle: " & objItem.BaseStyle
      WScript.Echo "BuiltIn: " & objItem.BuiltIn
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "NextStyle: " & objItem.NextStyle
      WScript.Echo "NoProofing: " & objItem.NoProofing
      WScript.Echo
   Next
Next

