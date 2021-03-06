' Description: Demonstration script that deletes all records from a database where the Department field is equal to Human Resources.


Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = inventory.mdb" 

objRecordSet.Open "DELETE * FROM GeneralProperties WHERE " & _
    "Department = 'Human Resources'", _
        objConnection, adOpenStatic, adLockOptimistic

objConnection.Close

