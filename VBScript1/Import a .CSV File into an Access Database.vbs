Const adOpenStatic = 3
Const adLockOptimistic = 3
Const ForReading = 1

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = c:\scripts\test.mdb" 

objRecordSet.Open "SELECT * FROM Employees", _
    objConnection, adOpenStatic, adLockOptimistic

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt")

Do Until objFile.AtEndOfStream
    strEmployee = objFile.ReadLine
    arrEmployee = Split(strEmployee, ",")

    objRecordSet.AddNew
    objRecordSet("EmployeeID") = arrEmployee(0)
    objRecordSet("EmployeeName") = arrEmployee(1)
    objRecordSet("Department") = arrEmployee(2)
    objRecordSet.Update

Loop

objRecordSet.Close
objConnection.Close
  


