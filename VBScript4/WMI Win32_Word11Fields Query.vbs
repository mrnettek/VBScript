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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Word11Fields", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "FieldCount: " & objItem.FieldCount
      WScript.Echo "ShowFieldCodes: " & objItem.ShowFieldCodes
      WScript.Echo "UpdateFieldsAtPrint: " & objItem.UpdateFieldsAtPrint
      WScript.Echo "UpdateLinksAtOpen: " & objItem.UpdateLinksAtOpen
      WScript.Echo
   Next
Next

