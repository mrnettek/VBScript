strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Payroll'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile in colFiles
    If objFile.Extension = "log" Then
        strCopy = "D:\Operation Logs\" & objFile.FileName _
            & "." & objFile.Extension
        objFile.Copy(strCopy)
        objFile.Delete
    End If
Next
  


