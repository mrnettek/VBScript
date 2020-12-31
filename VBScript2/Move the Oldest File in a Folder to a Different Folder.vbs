Set objFSo = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Upload")

Set colFiles = objFolder.Files

dtmOldestDate = Now
 
For Each objFile in colFiles
    If objFile.DateCreated < dtmOldestDate Then
        dtmOldestDate = objFile.DateCreated
        strOldestFile = objFile.Path
    End If
Next

objFSO.MoveFile strOldestFile, "C:\Download\"
  


