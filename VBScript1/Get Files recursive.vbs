
Call GetFiles("C:\Windows", True)

' ----------------------------------
Function GetFiles(sPath, bRecursive)

On Error Resume Next

Dim bCreateObject : bCreateObject = False

If not IsObject(osfo) then
	Dim ofso : Set ofso = CreateObject("Scripting.FileSystemObject")
	bCreateObject = True
End if

If ofso.FolderExists(sPath) = False then
	If bCreateObject and IsObject(osfo) then Set osfo = Nothing
	Exit Function
End if


	Dim oSearchPath : Set oSearchPath = ofso.Getfolder(sPath)
	Dim oSearchFiles : Set oSearchFiles = oSearchPath.Files

	If bRecursive then
		For Each oFolder in oSearchPath.Subfolders
			Call GetFiles(oFolder, bRecursive)
		Next
	End if

		For Each oFile in oSearchFiles
			' Do something with the files here
			Wscript.Echo oFile.Path
		Next

If IsObject(oSearchPath) then Set oSearchPath = Nothing
If IsObject(oSearchFiles) then Set oSearchFiles = Nothing
If bCreateObject and IsObject(osfo) then Set osfo = Nothing

End Function
