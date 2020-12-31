strOldestFile = ""
dtmOldestDate = Now

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\Scripts")

intFolderSize = Int((objFolder.Size / 1024) / 1024)

If intFolderSize >= 25 Then
    Set colFiles = objFolder.Files

    For Each objFile in colFiles
        strFile = objFile.Path
        dtmFileDate = objFile.DateCreated
        If dtmFileDate < dtmOldestDate Then
            dtmOldestDate = dtmFileDate
            strOldestFile = strFile
        End If
    Next

    objFSO.DeleteFile(strOldestFile) 
End If
  


