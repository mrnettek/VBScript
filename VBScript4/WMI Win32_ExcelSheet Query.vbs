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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ExcelSheet", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Cells: " & objItem.Cells
      WScript.Echo "Columns: " & objItem.Columns
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Rows: " & objItem.Rows
      WScript.Echo "Type: " & objItem.Type
      WScript.Echo
   Next
Next

