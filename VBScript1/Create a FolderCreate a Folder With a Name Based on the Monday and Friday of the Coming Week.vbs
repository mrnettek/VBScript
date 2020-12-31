intDay = Weekday(Date)

If intDay = 1 Then
    intAdder = 1
Else
    intAdder = 9 - intDay
End If

dtmMonday = Date + intAdder
dtmFriday = dtmMonday + 4

strFolderName = "C:\Test\" & dtmMonday & " to " & dtmFriday
strFolderName = Replace(strFolderName, "/", "-")

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateFolder(strFolderName)
  


