Const wdPageBreak = 7 

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set FileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Scripts\Archive'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile in FileList
    objSelection.InsertFile(objFile.Name)
    objSelection.InsertBreak(wdPageBreak)
Next
  


