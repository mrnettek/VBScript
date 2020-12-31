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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PowerPoint11PageSetup", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "BottomMargin: " & objItem.BottomMargin
      WScript.Echo "CharsLine: " & objItem.CharsLine
      WScript.Echo "FooterDistance: " & objItem.FooterDistance
      WScript.Echo "HeaderDistance: " & objItem.HeaderDistance
      WScript.Echo "LeftMargin: " & objItem.LeftMargin
      WScript.Echo "LinesPage: " & objItem.LinesPage
      WScript.Echo "Notables: " & objItem.Notables
      WScript.Echo "Orientation: " & objItem.Orientation
      WScript.Echo "PageHeight: " & objItem.PageHeight
      WScript.Echo "PageWidth: " & objItem.PageWidth
      WScript.Echo "PaperSize: " & objItem.PaperSize
      WScript.Echo "RightMargin: " & objItem.RightMargin
      WScript.Echo "Section: " & objItem.Section
      WScript.Echo "SectionStart: " & objItem.SectionStart
      WScript.Echo "TopMargin: " & objItem.TopMargin
      WScript.Echo "VerticalAlignment: " & objItem.VerticalAlignment
      WScript.Echo
   Next
Next

