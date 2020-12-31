Const ForReading = 1

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

objExcel.Workbooks.Add

i = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    arrLine = Split(strLine, ",")

    objExcel.Cells(i, 1).Value = arrLine(0)
    objExcel.Cells(i, 2).Value = arrLine(2)

    i = i + 1
Loop

objFile.Close
  


