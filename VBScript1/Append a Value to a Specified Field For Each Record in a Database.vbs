Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = c:\scripts\test.mdb" 

objRecordSet.Open "UPDATE TextFiles SET FileName = FileName & 's'", _
    objConnection, adOpenStatic, adLockOptimistic

objRecordSet.Close
objConnection.Close
  


