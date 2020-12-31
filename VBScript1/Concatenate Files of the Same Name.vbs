Const ForReading = 1
Const ForAppending = 8

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_Datafile Where FileName = 'test' AND Extension = 'vbs'")

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objMainFile = objFSO.OpenTextFile("C:\Scripts\Main_file.txt", ForAppending, True)

For Each objFile in colFiles
    Set objFile = objFSO.OpenTextFile(objFile.Name, ForReading)
    strContents = objFile.ReadAll
    objFile.Close
    objMainFile.WriteLine strContents
Next

objMainFile.Close
  


