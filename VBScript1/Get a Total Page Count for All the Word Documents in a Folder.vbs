Const wdStatisticPages = 2

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Scripts'} Where " _
        & "ResultClass = CIM_DataFile")

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

For Each objFile in colFiles
    If objFile.Extension = "doc" Then
        Set objDoc = objWord.Documents.Open(objFile.Name)
        intPages = intPages + objDoc.ComputeStatistics(wdStatisticPages)
        objDoc.Saved = True
        objDoc.Close
    End If
Next

Wscript.Echo "Total pages: " & intPages
objWord.Quit
  


