strMonth = Month(Date)

If Len(strMonth) = 1 Then
    strMonth = "0" & strMonth
End If

strDay = Day(Date)

If Len(strDay) = 1 Then
    strDay = "0" & strDay
End If

strYear = Year(Date)

strFolderName = "C:\Scripts\Tammy_" & strMonth & "-" & strDay & "-" & strYear

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.CreateFolder(strFolderName)
  


