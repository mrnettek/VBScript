Dim ofso 		: Set ofso		= CreateObject("Scripting.FileSystemObject")
Dim oShell 		: Set oShell 		= CreateObject("Shell.Application")
Dim sFile, sDir

If WScript.Arguments.Count > 0 then
	If WScript.Arguments.Count > 1 then
		For iCount = 0 to WScript.Arguments.Count -1
			If iCount = WScript.Arguments.Count -1 then
				sFile = sFile & WScript.Arguments(iCount)
			Else
				sFile = sFile & WScript.Arguments(iCount) & " "
			End if
		Next
	Else
		sFile = WScript.Arguments(0)
	End if
End if

If ofso.FileExists(sFile) then
	sDir = ofso.GetFile(sFile).ParentFolder
	sFile = ofso.GetFile(sFile).Name
	Dim oFolder : Set oFolder = oShell.Namespace(sDir)
	Dim oFolderItem : Set oFolderItem = oFolder.ParseName(sFile)

	Dim colVerbs : Set colVerbs = oFolderItem.Verbs

	For Each oVerb in colVerbs 

		sVerb = LCase(Replace(oVerb.Name, Chr(38), "", 1, -1, 1))

		Select Case sVerb
                    case "an startmenü anheften"
			oVerb.DoIt()
			WScript.Quit
                    case "pin to start menu"
			oVerb.DoIt()
			WScript.Quit
		End Select
	Next
End if

