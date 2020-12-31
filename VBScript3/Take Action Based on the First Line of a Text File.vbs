Const ForReading = 1

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Scripts'} Where " _
        & "ResultClass = CIM_DataFile")

Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each objItem In colItems
    Set objFile = objFSO.OpenTextFile(objItem.Name, ForReading)
    strLine = objFile.ReadLine
    strLine = LCase(strLine)
    objFile.Close

    If InStr(strLine, "fabrikam.com") or InStr(strLine, "contoso.com") Then
        objFSO.MoveFile objItem.Name, "C:\Archive\"
    End If 
Next
  


