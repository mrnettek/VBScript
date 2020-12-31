On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"

objRecordSet.Open "SELECT System.ItemPathDisplay, System.Photo.DateTaken FROM SYSTEMINDEX Where System.ItemFolderPathDisplay = 'C:\Test'", _
    objConnection

objRecordSet.MoveFirst

Do Until objRecordset.EOF
    strName = objRecordset.Fields.Item("System.ItemPathDisplay")
    arrName = Split(strName, ".")

    dtmPhotoDate = objRecordset.Fields.Item("System.Photo.DateTaken")
    strDay = Day(dtmPhotoDate)
    strMonth = Month(dtmPhotoDate)
    strYear = Year(dtmPhotoDate)
    strNewName = arrName(0) & "_" & strMonth & "_" & strDay & "_" & strYear & "." & arrName(1)
    
    objFSO.MoveFile strName , strNewName
    objRecordset.MoveNext
Loop
  


