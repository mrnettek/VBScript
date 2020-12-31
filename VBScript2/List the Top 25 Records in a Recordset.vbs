' Description: Demonstration script that queries a database for the 25 computers with the most physical memory.


Const adOpenStatic = 3
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = inventory.mdb" 

objRecordSet.Open "SELECT TOP 25 * FROM GeneralProperties " & _
    "ORDER BY TotalPhysicalMemory DESC", _
        objConnection, adOpenStatic, adLockOptimistic

objRecordSet.MoveFirst

Do Until objRecordset.EOF
    Wscript.Echo objRecordset.Fields.Item("ComputerName") & _
        vbTab & objRecordset.Fields.Item("TotalPhysicalMemory")
    objRecordset.MoveNext
Loop

objRecordSet.Close
objConnection.Close

