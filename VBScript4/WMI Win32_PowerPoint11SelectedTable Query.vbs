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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PowerPoint11SelectedTable", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AllowAutoFit: " & objItem.AllowAutoFit
      WScript.Echo "AllowPageBreaks: " & objItem.AllowPageBreaks
      WScript.Echo "BottomMargin: " & objItem.BottomMargin
      WScript.Echo "CellsAllowAutoFit: " & objItem.CellsAllowAutoFit
      WScript.Echo "Columns: " & objItem.Columns
      WScript.Echo "LeftIndent: " & objItem.LeftIndent
      WScript.Echo "LeftMargin: " & objItem.LeftMargin
      WScript.Echo "NestingLevel: " & objItem.NestingLevel
      WScript.Echo "Page: " & objItem.Page
      WScript.Echo "PreferredWidth: " & objItem.PreferredWidth
      WScript.Echo "PreferredWidthType: " & objItem.PreferredWidthType
      WScript.Echo "RightMargin: " & objItem.RightMargin
      WScript.Echo "RowAlignment: " & objItem.RowAlignment
      WScript.Echo "RowHeightRule: " & objItem.RowHeightRule
      WScript.Echo "Rows: " & objItem.Rows
      WScript.Echo "RowsAllowPageBreaks: " & objItem.RowsAllowPageBreaks
      WScript.Echo "Section: " & objItem.Section
      WScript.Echo "Spacing: " & objItem.Spacing
      WScript.Echo "Tables: " & objItem.Tables
      WScript.Echo "TopMargin: " & objItem.TopMargin
      WScript.Echo "Uniform: " & objItem.Uniform
      WScript.Echo
   Next
Next

