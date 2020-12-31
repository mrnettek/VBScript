strComputer = "atl-ds-01" 

Set objConnection = CreateObject("ADODB.Connection")

objConnection.Open _
    "Provider=SQLOLEDB;Data Source=" & strComputer & ";" & _
        "Trusted_Connection=Yes;Initial Catalog=Master"

Set objRecordset = objConnection.Execute("Select Name From SysDatabases")

If objRecordset.Recordcount = 0 Then
    Wscript.Echo "No databases could be found."
Else
    Do Until objRecordset.EOF
        Wscript.Echo objRecordset.Fields("Name")
        objRecordset.MoveNext
    Loop
End If
  


