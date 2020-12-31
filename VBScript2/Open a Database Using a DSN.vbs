' Description: Demonstration script that uses a DSN to open a database named Northwind.


Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Northwind;fabrikam\kenmyer;34ghfn&!j"

objRecordSet.Open "SELECT * FROM Customers", _
        objConnection, adOpenStatic, adLockOptimistic

objRecordSet.MoveFirst

Wscript.Echo objRecordSet.RecordCount

