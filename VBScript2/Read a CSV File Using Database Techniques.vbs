' Description: Demonstration script that uses ADO (Active Database Objects) to read a comma-separated-values text file.


On Error Resume Next

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

strPathtoTextFile = "C:\Databases\"

objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & strPathtoTextFile & ";" & _
          "Extended Properties=""text;HDR=YES;FMT=Delimited"""

objRecordset.Open "SELECT * FROM PhoneList.csv", _
          objConnection, adOpenStatic, adLockOptimistic, adCmdText

Do Until objRecordset.EOF
    Wscript.Echo "Name: " & objRecordset.Fields.Item("Name")
    Wscript.Echo "Department: " & _
        objRecordset.Fields.Item("Department")
    Wscript.Echo "Extension: " & objRecordset.Fields.Item("Extension")   
    objRecordset.MoveNext
Loop

