' Description: Uses the predefined query #AllProps to return the file name, file size, and author of all the files included in the Indexing Service catalog Script Catalog on the local computer.


On Error Resume Next

Set objConnection = CreateObject("ADODB.Connection")
objConnection.ConnectionString = "provider=msidxs;"
objConnection.Properties("Data Source") = "Script Catalog"
objConnection.Open
 
Set objCommand = CreateObject("ADODB.Command")
 
strQuery = "Create View #AllProps as Select * from Scope()"
 
Set objRecordSet = objConnection.Execute("Select * from Extended_FileInfo")
 
Do While Not objRecordSet.EOF
    Wscript.Echo objRecordSet("Filename") & ", " & objRecordSet("Size") & _
        ", " & objRecordSet("DocAuthor")
    objRecordSet.MoveNext
Loop

