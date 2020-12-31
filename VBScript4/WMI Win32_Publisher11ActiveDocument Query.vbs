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
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Publisher11ActiveDocument", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AllowFastSaves: " & objItem.AllowFastSaves
      WScript.Echo "Author: " & objItem.Author
      WScript.Echo "Chars: " & objItem.Chars
      WScript.Echo "CharsWithSpaces: " & objItem.CharsWithSpaces
      WScript.Echo "CreateDate: " & WMIDateStringToDate(objItem.CreateDate)
      WScript.Echo "LastAuthor: " & objItem.LastAuthor
      WScript.Echo "LastSavedDate: " & WMIDateStringToDate(objItem.LastSavedDate)
      WScript.Echo "Lines: " & objItem.Lines
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "Pages: " & objItem.Pages
      WScript.Echo "Paragraphs: " & objItem.Paragraphs
      WScript.Echo "Path: " & objItem.Path
      WScript.Echo "Sections: " & objItem.Sections
      WScript.Echo "Size: " & objItem.Size
      WScript.Echo "StoryTypes: " & objItem.StoryTypes
      WScript.Echo "Template: " & objItem.Template
      WScript.Echo "View: " & objItem.View
      WScript.Echo "WindowPosLeft: " & objItem.WindowPosLeft
      WScript.Echo "WindowPosTop: " & objItem.WindowPosTop
      WScript.Echo "Words: " & objItem.Words
      WScript.Echo "ZoomPercentage: " & objItem.ZoomPercentage
      WScript.Echo
   Next
Next


Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function

