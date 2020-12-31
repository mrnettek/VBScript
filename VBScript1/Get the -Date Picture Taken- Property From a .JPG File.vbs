Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"

objRecordSet.Open "SELECT System.FileName, System.Photo.DateTaken FROM SYSTEMINDEX " & _
    "Where System.ItemFolderPathDisplay = 'D:\Europe' and System.FileExtension = '.jpg'", _
    objConnection

objRecordSet.MoveFirst

Do Until objRecordset.EOF
    Wscript.Echo objRecordset.Fields.Item("System.FileName"), _
        objRecordset.Fields.Item("System.Photo.DateTaken")
    objRecordset.MoveNext
Loop
  


