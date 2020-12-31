' Description: Changes the file extension for all the .log files in the C:\Scripts folder to .txt.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set FileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='c:\Scripts'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In FileList
    If objFile.Extension = "log" Then
        strNewName = objFile.Drive & objFile.Path & _
            objFile.FileName & "." & "txt"
        errResult = objFile.Rename(strNewName)
    End If
Next

