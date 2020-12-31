
Call GetFolders("C:\Windows", True)

' ----------------------------------
Function GetFolders(sPath, bRecursive)

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

	If bRecursive then
		For Each oFolder in oSearchPath.Subfolders
			Call GetFolders(oFolder, bRecursive)
			' Do something with the folders here
			WScript.Echo oFolder.Path
		Next
	Else
		For Each oFolder in oSearchPath.Subfolders
			' Do something with the folders here
			WScript.Echo oFolder.Path
		Next
	End if

If IsObject(oSearchPath) then Set oSearchPath = Nothing
If bCreateObject and IsObject(osfo) then Set osfo = Nothing

End Function
