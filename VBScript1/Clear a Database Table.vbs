' Description: Demonstration script that deletes all the records in a table named Hardware from an ADO database with the DSN "Inventory."


Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordset = CreateObject("ADODB.Recordset")
objConnection.Open "DSN=Inventory;"
objRecordset.CursorLocation = adUseClient
objRecordset.Open "Delete * FROM Hardware" , objConnection, _
    adOpenStatic, adLockOptimistic
objConnection.Close

