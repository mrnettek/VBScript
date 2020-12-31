'*****************************************************
'************      Autor: Boris Toll      ************
'P: SCRENC					//////
'12:2004					//////
'File: toVBE.vbs				//////
'*****************************************************
' # Description:: Drag & drop the File to encode over the Script

If WScript.Arguments.Count = 0 Then
	WScript.Echo  "Kein Parameter angegeben"
Else
	On Error Resume Next

	Set fso = CreateObject("Scripting.FileSystemObject")

	For each Argument in WScript.Arguments
		skript = skript & Argument & " "
	Next

	Set Codex = fso.OpenTextFile(skript)
		code = Codex.ReadAll
		CHKerr()
		Codex.close

	Set SEncod = CreateObject("Scripting.Encoder")
		newcode = SEncod.EncodeScriptFile(".vbs", code, 0, "")

	script = fso.GetBaseName(skript)
	path = fso.GetParentFolderName(skript)

	newname = script & ".vbe"
	newpathname = fso.BuildPath(path, newname)

	Set newfile = fso.CreateTextFile(newpathname, true)
		newfile.Write newcode
		newfile.close

end if

Private Function CHKerr()

	if err.number <> 0 then
		if err.number = 62 then
			WScript.echo "Fehlercode: " & err.number & vbcrlf & err.description & vbcrlf & "Leere Dateien können nicht umgewandelt werden"
			err.clear
		else
			WScript.echo "Fehlercode: " & err.number & vbcrlf & err.description
			err.clear
		end if
	end if

End Function
