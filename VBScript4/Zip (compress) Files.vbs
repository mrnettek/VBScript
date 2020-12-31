'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'
'/|									  |\\\\\\\\'
'//|									   |\\\\\\\'
'///|									    |\\\\\\'
'////|			Version 	1.0.0				     |\\\\\'
'/////|			Author:		Boris TOll 			      |\\\\'
'//////|		Last Update:	15.05.2008			       |\\\'
'///////|								        |\\'
'////////|									 |\'
'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'
'
' # Description:: Drag & drop the Files to zip over the Script


Dim ofso		: Set ofso 	= Createobject("Scripting.FileSystemObject")
Dim oWshShell		: Set oWshShell = WScript.CreateObject("WSCript.shell")
Dim oApp    		: Set oApp 	= CreateObject("Shell.Application")

Const ForReading	= 1
Const ForWriting 	= 2
Const ForAppending 	= 8

Dim strZipTempPath	: strZipTempPath = oWshShell.ExpandEnvironmentStrings("%Temp%") & "\$tmpZip"


Call GetArguments(ArgArray)

If IsArray(ArgArray) then

	Dim strZipFile : strZipFile = SaveFile()

	If strZipFile > "" then

		If not UCase(Right(strZipFile, 4)) = ".ZIP" then
			strZipFile = strZipFile & ".zip"
		End if

		For Each ArrayElement In ArgArray

			If ofso.FileExists(ArrayElement) then

				Call Zip(ArrayElement, strZipFile)

			End if

		Next

		wscript.echo "Zipfile " & strZipFile & VbCrLf & "wurde erstellt!"

	Else
		wscript.echo "Script wurde abgebrochen!"
		wscript.quit
	End if


End if



' ---------------------------------------
Private Function Zip(strFile, strZipFile)

	If not ofso.FileExists(strZipFile) then
		Set oZip = ofso.OpenTextFile(strZipFile, ForWriting, True )
		oZip.Write "PK" & Chr(5) & Chr(6) & String( 18, Chr(0) )
		oZip.Close
		WScript.Sleep 2500
	End if

	tmpCount = oApp.NameSpace(strZipFile).Items.Count +1

	For Each Item in oApp.NameSpace(strZipFile).Items

		If UCase(Item) = Ucase(ofso.GetFile(strFile).Name) then
			tmpCount = oApp.NameSpace(strZipFile).Items.Count
			WScript.Sleep 5000
			Exit For
		End if

	Next

	oApp.NameSpace(strZipFile).CopyHere strFile

    	Do Until oApp.NameSpace(strZipFile).Items.Count = tmpCount
        	WScript.Sleep 500
   	Loop

End Function


' ----------------------------------------
Private Function GetArguments(SourceArray)

Dim iCount : iCount = 0

	If wscript.arguments.count > 0 then

		ReDim ArgArray(wscript.arguments.count -1)

		For Each Argument in wscript.arguments

			ArgArray(iCount) = Argument
			iCount = iCount +1
		Next


	iCount = Null
	GetArguments = ArgArray
		

	End if

End Function


' -------------------------
Private Function SaveFile()

	On Error Resume Next

	Dim ofso 		: Set ofso      	= CreateObject("Scripting.FileSystemObject")
	Dim oDlg 		: set oDlg 		= Wscript.CreateObject("MSComDlg.CommonDialog")
	Dim iRet

	If Err.Number <> 0 then

			Err.Clear
			Set oDlg  	= CreateObject("UserAccounts.CommonDialog")

			If Err.Number <> 0 then
				MsgBox "Notwendige Runtimes sind nicht vorhanden, Script wird beendet.",16 , "Info"
				WScript.Quit
			End if

	End if

  	oDlg.Filter = "Zip Files (*.zip)|*.zip"
  	oDlg.FilterIndex = 1
  	oDlg.MaxFileSize = 10000
  	oDlg.CancelError = true
  	oDlg.ShowSave

	If oDlg.Filename > "" then 

		If ofso.FileExists(oDlg.Filename) then
			iRet = MsgBox("Zipfile ist bereits vorhanden, wollen Sie die Dateien anhängen?",68 , "Info")
			If iRet = 6 then
				SaveFile = oDlg.Filename
			End if
		Else
			SaveFile = oDlg.Filename
		End if

	Else

		SaveFile = ""

	End if

End Function
