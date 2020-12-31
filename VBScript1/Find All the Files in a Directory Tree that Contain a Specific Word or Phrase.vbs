On Error Resume Next

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"

objRecordSet.Open "SELECT System.ItemPathDisplay FROM SYSTEMINDEX WHERE Contains('Alice') " & _
    "AND System.ItemPathDisplay LIKE 'C:\Scripts\%'", objConnection

objRecordSet.MoveFirst

Do Until objRecordset.EOF
    Wscript.Echo objRecordset.Fields.Item("System.ItemPathDisplay")
    objRecordset.MoveNext
Loop
  


