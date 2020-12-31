Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Logs'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In colFileList
    strFilePath = objFile.Name
    Set objTextFile = objFSO.OpenTextFile(strFilePath, ForReading)
    Do Until objTextFile.AtEndOfStream
        strLine = objTextFile.ReadLine
    Loop
    strMessage = strMessage & strLine & vbCrLf
    objTextFile.Close
Next

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()

Set objSelection = objWord.Selection
objSelection.TypeText strMessage
  


