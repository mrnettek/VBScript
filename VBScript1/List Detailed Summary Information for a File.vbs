' Description: Uses the Shell's Application object to retrieve detailed summary information including name, size, owner, and file attributes) for all the files in a folder.


Set objShell = CreateObject ("Shell.Application")
Set objFolder = objShell.Namespace ("C:\Scripts")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim arrHeaders(13)

For i = 0 to 13
    arrHeaders(i) = objFolder.GetDetailsOf (objFolder.Items, i)
Next

For Each strFileName in objFolder.Items
    For i = 0 to 13
        If i <> 9 then
            Wscript.echo arrHeaders(i) _
                & ": " & objFolder.GetDetailsOf (strFileName, i) 
        End If
    Next
    Wscript.Echo
Next

