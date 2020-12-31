' Description: Uses the Shell object to create a new folder named C:\Archive.


ParentFolder = "C:\" 

set objShell = CreateObject("Shell.Application")
set objFolder = objShell.NameSpace(ParentFolder) 
objFolder.NewFolder "Archive"

