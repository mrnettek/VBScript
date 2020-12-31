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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Word11Table", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Columns: " & objItem.Columns
      WScript.Echo "Index: " & objItem.Index
      WScript.Echo "NestedTables: " & objItem.NestedTables
      WScript.Echo "Page: " & objItem.Page
      WScript.Echo "Rows: " & objItem.Rows
      WScript.Echo "Section: " & objItem.Section
      WScript.Echo
   Next
Next

