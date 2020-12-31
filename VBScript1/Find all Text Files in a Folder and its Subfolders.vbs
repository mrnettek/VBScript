Dim ofso	: Set ofso 		= Createobject("Scripting.FileSystemObject")
Dim oWshShell	: Set oWshShell 	= WScript.CreateObject("WSCript.shell")


SearchPath = "C:\DATA"

GetFiles SearchPath


'''--------------------- Main Sub
Sub GetFiles(SearchPath)

Dim sPath, SearchFiles
Set sPath 		= ofso.Getfolder(SearchPath)
Set SearchFiles		= sPath.Files

		For Each Folder in sPath.Subfolders
			GetFiles(Folder)
		Next

		For Each File in SearchFiles
			
			If Ucase(Right(File.Name,4)) = ".TXT" then
				msgbox File
			End if
		Next

End Sub
