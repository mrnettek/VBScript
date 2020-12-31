' Description: Lists summary information for all the files in the folder C:\Scripts.


Const FILE_NAME = 0

Set objShell = CreateObject ("Shell.Application")
Set objFolder = objShell.Namespace ("C:\Scripts")

For Each strFileName in objFolder.Items
    Wscript.Echo "File name: " & objFolder.GetDetailsOf _
        (strFileName, FILE_NAME) 
Next

