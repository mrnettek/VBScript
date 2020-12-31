' Description: Demonstration script that sets the value of the Department field to Accounting for all the records in a table.


Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = inventory.mdb" 

objRecordSet.Open "UPDATE GeneralProperties SET " & _
    "Department = 'Accounting'", _
        objConnection, adOpenStatic, adLockOptimistic

objConnection.Close

