' Description: Demonstration script that returns the unique operating systems found in a database.


Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = inventory.mdb" 

objRecordSet.Open "SELECT DISTINCT OSName FROM " & _
    "GeneralProperties ORDER BY OSName", _
        objConnection, adOpenStatic, adLockOptimistic

objRecordSet.MoveFirst

Do Until objRecordset.EOF
    Wscript.Echo objRecordset.Fields.Item("OSName")
    objRecordset.MoveNext
Loop

objRecordSet.Close
objConnection.Close

