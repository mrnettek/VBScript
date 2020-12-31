'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'
'/|									  |\\\\\\\\'
'//|									   |\\\\\\\'
'///|			#-------------#					    |\\\\\\'
'////|			Version 	1.0.0.1				     |\\\\\'
'/////|			Boris TOll 	15.03.2010			      |\\\\'
'//////|		Last Update:	15.03.2010			       |\\\'
'///////|								        |\\'
'////////|									 |\'
'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'


On Error Resume Next

Dim oWMIService 	: Set oWMIService 	= GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Dim oWshNetwork 	: Set oWshNetwork 	= WScript.CreateObject("WScript.Network")
Dim colItems 		: Set colItems		= oWMIService.ExecQuery("Select * from Win32_MappedLogicalDisk")
Dim colShares 		: Set colShares 	= oWMIService.ExecQuery("Select * from Win32_Share")
Dim ofso		: Set ofso 		= CreateObject("Scripting.FileSystemObject")
Dim oWshShell		: Set oWshShell		= CreateObject("WScript.Shell")
Dim Parameter		: Parameter 		= ""

Dim oWord 		: Set oWord 		= CreateObject("Word.Application")
If Err.Number <> 0 then
	MsgBox "Microsoft Word is not installed!",16,"Error"
Else
	oWord.Documents.Add
End if



'--- Drag and drop file or folder over the script
If WScript.Arguments.Count = 0 Then
	WScript.Echo  "Arguments required!"
	oWord.ActiveDocument.Close(0)
	oWord.Quit()
	wscript.quit
Else

	For each Arg in WScript.Arguments
		Parameter = Parameter & Arg & ";"
	Next

	Parameter = Left(Parameter,Len(Parameter) -1)

End if


If ofso.FileExists(Parameter) then

	For Each oItem in colItems
		If UCase(Left(oItem.DeviceID,2)) = UCase(Left(Parameter,2)) then
			Parameter = Replace(Parameter, UCase(Left(Parameter,2)), oItem.ProviderName, 1, -1, 1)
		End if		
	Next

	For each oShare in colShares
		If Len(Trim(oShare.Path)) > 0 and ofso.FolderExists(oShare.Path) = True and UCase(Left(oShare.Path,Len(Trim(oShare.Path)))) = UCase(Left(Parameter,Len(Trim(oShare.Path)))) then
			Parameter = Replace(Parameter, Left(Parameter,Len(Trim(oShare.Path))), "\\" & oWshNetwork.ComputerName & "\" & oShare.Name & "\", 1, -1, 1)
		End if
	Next


	Parameter = Replace(Parameter, "\", "/", 1, -1, 1)
	Parameter = Replace(Parameter, " ", "%20", 1, -1, 1)
	Parameter = Replace(Parameter, "ä", "%E4", 1, -1, 1)
	Parameter = Replace(Parameter, "Ä", "%C4", 1, -1, 1)
	Parameter = Replace(Parameter, "ü", "%FC", 1, -1, 1)
	Parameter = Replace(Parameter, "Ü", "%DC", 1, -1, 1)
	Parameter = Replace(Parameter, "ö", "%F6", 1, -1, 1)
	Parameter = Replace(Parameter, "Ö", "%D6", 1, -1, 1)
	Parameter = Replace(Parameter, "ß", "%DF", 1, -1, 1)
	Parameter = "file:///" & Parameter

End if

If ofso.FolderExists(Parameter) then

	For Each oItem in colItems
		If UCase(Left(oItem.DeviceID,2)) = UCase(Left(Parameter,2)) then
			Parameter = Replace(Parameter, UCase(Left(Parameter,2)), oItem.ProviderName, 1, -1, 1)
		End if		
	Next

	For each oShare in colShares
		If Len(Trim(oShare.Path)) > 0 and ofso.FolderExists(oShare.Path) = True and UCase(Left(oShare.Path,Len(Trim(oShare.Path)))) = UCase(Left(Parameter,Len(Trim(oShare.Path)))) then
			Parameter = Replace(Parameter, Left(Parameter,Len(Trim(oShare.Path))), "\\" & oWshNetwork.ComputerName & "\" & oShare.Name & "\", 1, -1, 1)
		End if
	Next

	Parameter = Replace(Parameter, "\", "/", 1, -1, 1)
	Parameter = Replace(Parameter, " ", "%20", 1, -1, 1)
	Parameter = Replace(Parameter, "ä", "%E4", 1, -1, 1)
	Parameter = Replace(Parameter, "Ä", "%C4", 1, -1, 1)
	Parameter = Replace(Parameter, "ü", "%FC", 1, -1, 1)
	Parameter = Replace(Parameter, "Ü", "%DC", 1, -1, 1)
	Parameter = Replace(Parameter, "ö", "%F6", 1, -1, 1)
	Parameter = Replace(Parameter, "Ö", "%D6", 1, -1, 1)
	Parameter = Replace(Parameter, "ß", "%DF", 1, -1, 1)
	Parameter = "file:///" & Parameter & "/"

End if

If Len(Parameter) > 0 then
	oword.selection.TypeText(Parameter)
	oWord.Selection.WholeStory()
	oWord.Selection.Copy()
	oWord.ActiveDocument.Close(0)
	oWord.Quit()
	Call oWshShell.Popup("Hyperlink was copied",1,"Copy as Hyperlink",0) 
Else
	oWord.ActiveDocument.Close(0)
	oWord.Quit()
End if
