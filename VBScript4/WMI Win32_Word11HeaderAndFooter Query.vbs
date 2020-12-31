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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Word11HeaderAndFooter", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "EvenFooterChars: " & objItem.EvenFooterChars
      WScript.Echo "EvenFooterLinkedToPrevious: " & objItem.EvenFooterLinkedToPrevious
      WScript.Echo "EvenFooterText: " & objItem.EvenFooterText
      WScript.Echo "EvenHeaderChars: " & objItem.EvenHeaderChars
      WScript.Echo "EvenHeaderLinkedToPrevious: " & objItem.EvenHeaderLinkedToPrevious
      WScript.Echo "EvenHeaderText: " & objItem.EvenHeaderText
      WScript.Echo "FirstFooterChars: " & objItem.FirstFooterChars
      WScript.Echo "FirstFooterLinkedToPrevious: " & objItem.FirstFooterLinkedToPrevious
      WScript.Echo "FirstFooterText: " & objItem.FirstFooterText
      WScript.Echo "FirstHeaderChars: " & objItem.FirstHeaderChars
      WScript.Echo "FirstHeaderLinkedToPrevious: " & objItem.FirstHeaderLinkedToPrevious
      WScript.Echo "FirstHeaderText: " & objItem.FirstHeaderText
      WScript.Echo "FooterChars: " & objItem.FooterChars
      WScript.Echo "FooterDistance: " & objItem.FooterDistance
      WScript.Echo "FooterLinkedToPrevious: " & objItem.FooterLinkedToPrevious
      WScript.Echo "FooterText: " & objItem.FooterText
      WScript.Echo "HeaderChars: " & objItem.HeaderChars
      WScript.Echo "HeaderDistance: " & objItem.HeaderDistance
      WScript.Echo "HeaderLinkedToPrevious: " & objItem.HeaderLinkedToPrevious
      WScript.Echo "HeaderText: " & objItem.HeaderText
      WScript.Echo "Notables: " & objItem.Notables
      WScript.Echo "Section: " & objItem.Section
      WScript.Echo
   Next
Next

