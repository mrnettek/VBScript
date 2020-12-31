' Description: Demonstration script that sorts a recordset on two fields: first by EventCode (in ascending order), and then by Logfile (in descending order).


Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = eventlogs.mdb" 

objRecordSet.Open "SELECT * FROM EventTable " & _
    "ORDER BY EventCode ASC, Logfile DESC", _
        objConnection, adOpenStatic, adLockOptimistic

objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    Wscript.Echo objRecordSet.Fields.Item("EventCode"), objRecordSet.Fields.Item("Logfile")
    objRecordSet.MoveNext
Loop

objRecordSet.Close
objConnection.Close
