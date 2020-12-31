'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'
'/|									  |\\\\\\\\'
'//|									   |\\\\\\\'
'///|									    |\\\\\\'
'////|			Version 	1.0.0				     |\\\\\'
'/////|			Author:		Boris TOll 			      |\\\\'
'//////|		Last Update:	31.01.2008			       |\\\'
'///////|								        |\\'
'////////|									 |\'
'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'


Dim strFile : strFile = OpenFile()

If strFile > "" then

	Msgbox "Ausgewählte Datei: " & VbCrLf & strFile

End if


'--------------------------------------------------------
Function OpenFile()

	On Error Resume Next

	Dim ofso 		: Set ofso      	= CreateObject("Scripting.FileSystemObject")
	Dim oDlg 		: set oDlg 		= Wscript.CreateObject("MSComDlg.CommonDialog")

	If Err.Number <> 0 then

			Err.Clear
			Set oDlg  	= CreateObject("UserAccounts.CommonDialog")

			If Err.Number <> 0 then
				MsgBox "Notwendige Runtimes sind nicht vorhanden, Script wird beendet.",16 , "Info"
				WScript.Quit
			End if

	End if

  	oDlg.Filter = "All Files (*.*)|*.*"
  	oDlg.FilterIndex = 1
  	oDlg.MaxFileSize = 10000
  	oDlg.CancelError = true
  	oDlg.ShowOpen

	If oDlg.Filename > "" and ofso.FileExists(oDlg.Filename) then

		OpenFile = oDlg.Filename

	Else

		OpenFile = ""

	End if

End Function
