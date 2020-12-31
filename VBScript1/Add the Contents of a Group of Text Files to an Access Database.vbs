Const ForReading = 1
Const adLockOptimistic = 3

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

Set objFSO = CreateObject("Scripting.FileSystemObject")

objConnection.Open _
    "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        "Data Source = C:\Scripts\Test.mdb" 

objRecordSet.Open "SELECT * FROM TextFiles" , _
    objConnection, adOpenStatic, adLockOptimistic

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Archive'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In colFileList
    Set objTextFile = objFSO.OpenTextFile(objFile.Name, ForReading)
    strContents = objTextFile.ReadAll
    objTextFile.Close

    objRecordSet.AddNew
    objRecordSet("FileName") = objFile.Name
    objRecordSet("FileContents") = strContents
    objRecordSet.Update
Next

objRecordSet.Close
objConnection.Close
  


