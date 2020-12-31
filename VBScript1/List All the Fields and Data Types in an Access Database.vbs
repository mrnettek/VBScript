Const adSchemaTables = 20
Const adSchemaColumns = 4

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = 'C:\Scripts\Test.mdb'" 

Set objRecordSet = objConnection.OpenSchema(adSchemaTables)

Do Until objRecordset.EOF
    strTableName = objRecordset("Table_Name")
    Set objFieldSchema = objConnection.OpenSchema(adSchemaColumns, _
        Array(Null, Null, strTableName))
    Wscript.Echo UCase(objRecordset("Table_Name"))

    Do While Not objFieldSchema.EOF
        Wscript.Echo objFieldSchema("Column_Name") & ", " & objFieldSchema("Data_Type")
        objFieldSchema.MoveNext
    Loop

    Wscript.Echo
    objRecordset.MoveNext
Loop
  


