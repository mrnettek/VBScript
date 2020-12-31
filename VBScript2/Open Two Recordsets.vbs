' Description: Demonstration script that works with two separate recordsets.


Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")
Set objRecordSet2 = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider= Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source=inventory.mdb" 

objRecordSet.Open "SELECT * FROM GeneralProperties Where ComputerName = 'Computer1'", _
        objConnection, adOpenStatic, adLockOptimistic

objRecordSet.MoveFirst


objRecordSet2.Open "SELECT * FROM Storage Where ComputerName = 'Computer1'", _
        objConnection, adOpenStatic, adLockOptimistic

objRecordSet2.MoveFirst

Do Until objRecordset.EOF
    Wscript.Echo objRecordset.Fields.Item("ComputerName")
    Wscript.Echo objRecordset.Fields.Item("OSName")
    objRecordSet.MoveNext
Loop

Do Until objRecordset2.EOF
    Wscript.Echo objRecordset2.Fields.Item("DriveName"), _
        objRecordset2.Fields.Item("DriveDescription")
    objRecordSet2.MoveNext
Loop

objRecordSet.Close
objRecordSet2.Close
objConnection.Close

